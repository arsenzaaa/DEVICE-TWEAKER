"""Prepare or replace the mutable 0.0.4 release. Preparation never writes to GitHub."""
import argparse
import datetime as dt
import hashlib
import json
from pathlib import Path
import posixpath
import re
import subprocess
import urllib.parse
import uuid
import zipfile

TAG = 'v0.0.4'
FILES = ('DEVICE.TWEAKER.exe', 'DEVICE.TWEAKER.SelfContained.exe',
         'DEVICE.TWEAKER.Source.v0.0.4.zip', 'SHA256SUMS.txt', 'SOURCE-SHA256SUMS.txt')


def run(*args, data=None):
    result = subprocess.run(args, input=data.encode('utf-8') if data is not None else None,
                            stdout=subprocess.PIPE, stderr=subprocess.PIPE)
    if result.returncode:
        message = result.stderr.decode('utf-8', errors='replace').strip()
        raise RuntimeError(f'{args[0]} {args[1]} failed: {message}')
    return result.stdout.decode('utf-8').strip()


def sha(path):
    with path.open('rb') as stream:
        return hashlib.file_digest(stream, 'sha256').hexdigest()


def save(path, value):
    path.write_text(json.dumps(value, ensure_ascii=False, indent=2) + '\n', encoding='utf-8')


def package_hashes(package):
    if sorted(p.name for p in package.iterdir()) != sorted(FILES):
        raise ValueError('Package must contain exactly the five release files')
    if any((package / name).is_symlink() or not (package / name).is_file() for name in FILES):
        raise ValueError('Package files must be regular files')
    hashes = {name: sha(package / name) for name in FILES}
    for manifest, expected in [('SHA256SUMS.txt', set(FILES[:2])),
                               ('SOURCE-SHA256SUMS.txt', {FILES[2]})]:
        found = {}
        for line in (package / manifest).read_text(encoding='utf-8-sig').splitlines():
            match = re.fullmatch(r'([0-9a-fA-F]{64})  ([^/\\]+)', line)
            if not match or match[2] in found:
                raise ValueError(f'Invalid checksum manifest: {manifest}')
            found[match[2]] = match[1].lower()
        if set(found) != expected or any(hashes[name] != value for name, value in found.items()):
            raise ValueError(f'Checksum mismatch: {manifest}')
    return hashes


def verify_sources(package, source):
    listing = run('git', 'ls-tree', '-rz', source)
    tree = {}
    for entry in listing.split('\0'):
        if entry:
            info, name = entry.split('\t', 1)
            mode, kind, oid = info.split()
            if kind != 'blob':
                raise ValueError('Source archive cannot contain submodules')
            tree[name] = oid
    with zipfile.ZipFile(package / FILES[2]) as archive:
        entries = [entry for entry in archive.infolist() if not entry.is_dir()]
        if archive.comment.decode('ascii') != source or sorted(e.filename for e in entries) != sorted(tree):
            raise ValueError('Source archive does not match the selected commit')
        for entry in entries:
            data = archive.read(entry)
            oid = hashlib.sha1(b'blob ' + str(len(data)).encode('ascii') + b'\0' + data).hexdigest()
            if tree[entry.filename] != oid:
                raise ValueError(f'Source archive content differs: {entry.filename}')


def release_notes(text, repository, source):
    """Resolve repository links before using a document as a GitHub release body."""
    def resolve(match):
        url = urllib.parse.urlsplit(match[2])
        if url.scheme or url.netloc or not url.path:
            return match[0]
        path = posixpath.normpath(posixpath.join('docs/releases', urllib.parse.unquote(url.path)))
        if path.startswith('../') or path.startswith('/'):
            raise ValueError('Release note link escapes the source tree')
        target = f'https://github.com/{repository}/blob/{source}/' + urllib.parse.quote(path, safe='/')
        if url.query:
            target += '?' + url.query
        if url.fragment:
            target += '#' + url.fragment
        return match[1] + target + match[3]
    return re.sub(r'(\[[^\]\n]+\]\()([^)\n]+)(\))', resolve, text)


class Release:
    def __init__(self, args):
        self.args = args
        self.base = f'repos/{args.repository}'
        self.backup = Path(args.backup).resolve()
        self.package = Path(args.package).resolve()

    def api(self, route, method='GET', payload=None):
        command = ['gh', 'api', f'{self.base}/{route}', '--method', method]
        if payload is not None:
            command += ['--input', '-']
        text = run(*command, data=json.dumps(payload) if payload is not None else None)
        return json.loads(text) if text else None

    def state(self):
        return {'main': self.api('git/ref/heads/main')['object']['sha'],
                'tag': self.api(f'git/ref/tags/{TAG}')['object'],
                'release': self.api(f'releases/tags/{TAG}')}

    @staticmethod
    def fingerprint(state):
        release = state['release']
        keys = ('id', 'tag_name', 'target_commitish', 'name', 'body', 'draft',
                'prerelease', 'immutable', 'updated_at', 'html_url')
        result = {key: release.get(key) for key in keys}
        result['assets'] = sorted([{key: asset.get(key) for key in
                                   ('id', 'name', 'label', 'size', 'digest', 'state', 'updated_at')}
                                  for asset in release['assets']], key=lambda a: a['name'])
        return {'main': state['main'], 'tag': state['tag'], 'release': result}

    def validate_state(self, state):
        args = self.args
        release = state['release']
        if state['main'] != args.expected_main or state['tag']['sha'] != args.expected_tag_object:
            raise ValueError('Remote main or tag changed; review the new state before replacing')
        if release['updated_at'] != args.expected_release_updated_at:
            raise ValueError('Release metadata changed; review before replacing')
        if release.get('immutable') or release['draft'] or release['prerelease']:
            raise ValueError('Expected the existing public, mutable stable release')
        if release['tag_name'] != TAG or sorted(a['name'] for a in release['assets']) != sorted(FILES):
            raise ValueError('Unexpected release tag or asset set')

    def download(self, directory, expected):
        directory.mkdir(parents=True, exist_ok=False)
        run('gh', 'release', 'download', TAG, '--repo', self.args.repository, '--dir', str(directory))
        actual = package_hashes(directory)
        if actual != expected:
            raise ValueError('Downloaded release differs from the verified file hashes')

    def prepare(self):
        candidate = package_hashes(self.package)
        if candidate['SHA256SUMS.txt'] != self.args.expected_exe_sums_sha256:
            raise ValueError('EXE checksums differ from the package accepted for publication')
        if self.args.source_sha != self.args.expected_main:
            raise ValueError('Publish the reviewed candidate to main before replacing the release')
        if run('git', 'rev-parse', 'HEAD') != self.args.source_sha:
            raise ValueError('Checkout differs from the selected source SHA')
        verify_sources(self.package, self.args.source_sha)
        state = self.state()
        self.validate_state(state)
        self.backup.mkdir(parents=True, exist_ok=False)
        save(self.backup / 'remote-before.json', state)
        tag = state['tag']
        if tag['type'] == 'tag':
            save(self.backup / 'tag-object.json', self.api(f"git/tags/{tag['sha']}"))
        # Keep the old annotated tag and its source locally as well as in API JSON.
        run('git', 'cat-file', '-e', tag['sha'])
        run('git', 'bundle', 'create', str(self.backup / 'before.bundle'), f'refs/tags/{TAG}')
        expected = {}
        for asset in state['release']['assets']:
            if not re.fullmatch(r'sha256:[0-9a-f]{64}', asset.get('digest') or ''):
                raise ValueError('A release asset has no verifiable SHA-256 digest')
            expected[asset['name']] = asset['digest'][7:]
        self.download(self.backup / 'assets', expected)
        if self.fingerprint(self.state()) != self.fingerprint(state):
            raise ValueError('Remote state changed during backup; nothing was replaced')
        notes = release_notes(Path('docs/releases/v0.0.4.md').read_text(encoding='utf-8-sig').rstrip(),
                              self.args.repository, self.args.source_sha)
        notes += f'\n\nSource commit: [{self.args.source_sha}](https://github.com/{self.args.repository}/tree/{self.args.source_sha}).\n'
        save(self.backup / 'prepared.json', {'source': self.args.source_sha, 'repository': self.args.repository,
             'package': candidate, 'previous': expected, 'notes': notes, 'state': self.fingerprint(state)})
        print('PREPARED: old release archived and verified; GitHub is unchanged.')

    def push_tag(self, target, expected):
        # A tag update is conditional on its actual old object, including annotated tags.
        run('git', 'push', 'origin', f'{target}:refs/tags/{TAG}',
            f'--force-with-lease=refs/tags/{TAG}:{expected}')

    def upload(self, directory):
        run('gh', 'release', 'upload', TAG, *(str(directory / name) for name in FILES),
            '--repo', self.args.repository, '--clobber')

    def restore(self, previous, new_tag):
        current = self.state()
        if current['main'] != previous['main'] or current['tag']['sha'] not in (previous['tag']['sha'], new_tag):
            raise RuntimeError('Concurrent remote change; automatic rollback stopped. Keep the release draft and use the backup.')
        release = previous['release']
        route = f"releases/{release['id']}"
        self.api(route, 'PATCH', {'draft': True})
        if current['tag']['sha'] != previous['tag']['sha']:
            self.push_tag(previous['tag']['sha'], new_tag)
        expected = json.loads((self.backup / 'prepared.json').read_text(encoding='utf-8'))['previous']
        assets = current['release']['assets']
        current_hashes = {a['name']: (a.get('digest') or '').removeprefix('sha256:') for a in assets}
        if len(assets) != len(FILES) or current_hashes != expected or any(a.get('state') != 'uploaded' for a in assets):
            self.upload(self.backup / 'assets')
        # clobber can leave a failed "starter" upload. Remove only unexpected assets in our draft.
        current_assets = self.api(route)['assets']
        for asset in current_assets:
            if asset['name'] not in FILES or asset.get('state') == 'starter':
                self.api(f"releases/assets/{asset['id']}", 'DELETE')
            else:
                original = next(a for a in release['assets'] if a['name'] == asset['name'])
                if original.get('label'):
                    self.api(f"releases/assets/{asset['id']}", 'PATCH', {'label': original['label']})
        self.download(self.backup / ('restored-download-' + uuid.uuid4().hex), expected)
        self.api(route, 'PATCH', {key: release[key] for key in
                                 ('name', 'body', 'tag_name', 'target_commitish', 'draft', 'prerelease')})
        save(self.backup / 'rollback.json', {'status': 'RESTORED', 'state': self.state()})

    def recover(self):
        prepared = json.loads((self.backup / 'prepared.json').read_text(encoding='utf-8'))
        previous = json.loads((self.backup / 'remote-before.json').read_text(encoding='utf-8'))
        tag = json.loads((self.backup / 'replacement-tag.json').read_text(encoding='utf-8'))
        if prepared['repository'] != self.args.repository or prepared['source'] != self.args.source_sha:
            raise ValueError('Recovery archive belongs to another repository or source')
        if self.fingerprint(previous) != prepared['state']:
            raise ValueError('Archived release metadata changed; no remote writes made')
        if package_hashes(self.backup / 'assets') != prepared['previous']:
            raise ValueError('Recovery archive changed; no remote writes made')
        bundle = str(self.backup / 'before.bundle')
        run('git', 'bundle', 'verify', bundle)
        heads = run('git', 'bundle', 'list-heads', bundle, f'refs/tags/{TAG}').split()
        if heads != [previous['tag']['sha'], f'refs/tags/{TAG}']:
            raise ValueError('Archived bundle does not contain the expected old tag')
        run('git', 'fetch', bundle, f'refs/tags/{TAG}')
        self.restore(previous, tag['sha'])
        print('RESTORED: archived release files, tag and description verified.')

    def apply(self):
        prepared = json.loads((self.backup / 'prepared.json').read_text(encoding='utf-8'))
        previous = json.loads((self.backup / 'remote-before.json').read_text(encoding='utf-8'))
        if prepared['source'] != self.args.source_sha or prepared['repository'] != self.args.repository:
            raise ValueError('Backup belongs to another source or repository')
        if self.fingerprint(previous) != prepared['state']:
            raise ValueError('Archived release metadata changed; no remote writes made')
        if prepared['package']['SHA256SUMS.txt'] != self.args.expected_exe_sums_sha256:
            raise ValueError('EXE checksums differ from the accepted package')
        if package_hashes(self.package) != prepared['package'] or package_hashes(self.backup / 'assets') != prepared['previous']:
            raise ValueError('Candidate or backup changed after preparation')
        current = self.state()
        self.validate_state(current)
        if self.fingerprint(current) != prepared['state']:
            raise ValueError('Release changed after preparation; nothing was replaced')
        if run('git', 'rev-parse', 'HEAD') != self.args.source_sha:
            raise ValueError('Checkout changed after preparation')
        timestamp = int(dt.datetime.now(dt.timezone.utc).timestamp())
        tag_text = (f'object {self.args.source_sha}\ntype commit\ntag {TAG}\n'
                    f'tagger DEVICE TWEAKER release <41898282+github-actions[bot]@users.noreply.github.com> {timestamp} +0000\n'
                    f'\nDEVICE TWEAKER 0.0.4\n')
        new_tag = run('git', 'mktag', data=tag_text)
        save(self.backup / 'replacement-tag.json', {'sha': new_tag, 'source': self.args.source_sha})
        route = f"releases/{previous['release']['id']}"
        try:
            self.api(route, 'PATCH', {'draft': True})
            self.push_tag(new_tag, previous['tag']['sha'])
            self.upload(self.package)
            self.download(self.backup / 'replacement-download', prepared['package'])
            current = self.state()
            if current['main'] != self.args.expected_main or current['tag']['sha'] != new_tag:
                raise ValueError('Remote refs changed during replacement')
            if current['release']['id'] != previous['release']['id']:
                raise ValueError('Release identity changed during replacement')
            notes = prepared['notes'] + f'\nUpdated: {dt.datetime.now(dt.timezone.utc).date().isoformat()} (UTC).\n'
            self.api(route, 'PATCH', {'name': 'DEVICE TWEAKER 0.0.4', 'body': notes,
                                     'target_commitish': self.args.source_sha, 'draft': False, 'prerelease': False})
            self.download(self.backup / 'published-download', prepared['package'])
            published = self.state()
            if published['release']['html_url'] != previous['release']['html_url'] or published['release']['draft']:
                raise ValueError('Published release URL or visibility differs from the expected release')
            save(self.backup / 'replacement.json', {'status': 'VERIFIED', 'state': published})
        except Exception as failure:
            try:
                self.restore(previous, new_tag)
            except Exception as recovery_failure:
                save(self.backup / 'rollback-error.json', {'failure': str(failure), 'recovery': str(recovery_failure)})
                raise RuntimeError('Replacement and recovery failed; inspect the archived state before any further write') from recovery_failure
            raise RuntimeError('Replacement failed; the previous release was restored') from failure
        print('VERIFIED: existing release replaced; published downloads match the package.')


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--repository', default='arsenzaaa/DEVICE-TWEAKER')
    parser.add_argument('--source-sha', required=True)
    parser.add_argument('--expected-main', required=True)
    parser.add_argument('--expected-tag-object', required=True)
    parser.add_argument('--expected-release-updated-at', required=True)
    parser.add_argument('--expected-exe-sums-sha256', required=True)
    parser.add_argument('--package', required=True)
    parser.add_argument('--backup', required=True)
    mode = parser.add_mutually_exclusive_group()
    mode.add_argument('--apply', action='store_true', help='Replace remote release using an existing verified backup')
    mode.add_argument('--restore', action='store_true', help='Recover an interrupted replacement from its verified backup')
    args = parser.parse_args()
    if not re.fullmatch(r'[A-Za-z0-9_.-]+/[A-Za-z0-9_.-]+', args.repository):
        parser.error('Invalid repository')
    for value in (args.source_sha, args.expected_main, args.expected_tag_object):
        if not re.fullmatch('[0-9a-f]{40}', value):
            parser.error('Use full 40-character commit/object SHA values')
    if not re.fullmatch('[0-9a-f]{64}', args.expected_exe_sums_sha256):
        parser.error('Use the SHA-256 of SHA256SUMS.txt from the accepted package')
    release = Release(args)
    if args.restore:
        release.recover()
    elif args.apply:
        release.apply()
    else:
        release.prepare()


if __name__ == '__main__':
    main()
