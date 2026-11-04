"""Offline failure-injection checks. No GitHub requests or release writes are made."""
import copy
import hashlib
import json
from pathlib import Path
import shutil
import tempfile
import types
import unittest
from unittest.mock import patch
import zipfile

import replace_release as release

SOURCE = 'a' * 40
OLD = 'b' * 40
NEW = 'c' * 40


def make_package(path, content):
    path.mkdir()
    for name in release.FILES[:3]:
        (path / name).write_bytes(content + name.encode())
    for manifest, names in [('SHA256SUMS.txt', release.FILES[:2]),
                            ('SOURCE-SHA256SUMS.txt', [release.FILES[2]])]:
        (path / manifest).write_text(''.join(f'{release.sha(path / name)}  {name}\n' for name in names), encoding='ascii')
    return release.package_hashes(path)


class Remote:
    def __init__(self, args, old_package, old_hashes):
        self.args = args
        self.old_package = old_package
        self.hashes = dict(old_hashes)
        self.tag = OLD
        self.main = SOURCE
        self.metadata = {'id': 42, 'tag_name': release.TAG, 'target_commitish': 'old',
                         'name': 'Original title', 'body': 'Original notes', 'draft': False,
                         'prerelease': False, 'immutable': False, 'updated_at': 'before',
                         'html_url': 'https://github.com/example/project/releases/tag/v0.0.4'}
        self.writes = []
        self.fail_at = None
        self.fail_used = False

    def snapshot(self):
        data = copy.deepcopy(self.metadata)
        if data['draft']:
            data['html_url'] = 'https://github.com/example/project/releases/tag/untagged-draft'
        data['assets'] = [{'id': i, 'name': name, 'label': '', 'size': 1,
                           'digest': 'sha256:' + value, 'state': 'uploaded', 'updated_at': 'before'}
                          for i, (name, value) in enumerate(self.hashes.items())]
        return {'main': self.main, 'tag': {'sha': self.tag, 'type': 'tag'}, 'release': data}

    def fail(self, point):
        if self.fail_at == point and not self.fail_used:
            self.fail_used = True
            raise RuntimeError('Injected ' + point)

    def api(self, route, method='GET', payload=None):
        if route == 'git/ref/heads/main':
            return {'object': {'sha': self.main}}
        if route == f'git/ref/tags/{release.TAG}':
            return {'object': {'sha': self.tag, 'type': 'tag'}}
        if route == f'releases/tags/{release.TAG}':
            if self.metadata['draft']:
                raise RuntimeError('gh: Not Found (HTTP 404): draft releases are not returned by tag')
            return self.snapshot()['release']
        if route != 'releases/42':
            raise AssertionError(route)
        if method == 'PATCH':
            self.writes.append(('metadata', copy.deepcopy(payload)))
            self.metadata.update(payload)
            if payload.get('draft') is True:
                self.fail('draft-response')
            if payload.get('draft') is False:
                self.fail('publish-response')
        if route.startswith('releases/'):
            return self.snapshot()['release']
        raise AssertionError(route)

    def push_tag(self, target, expected):
        if self.tag != expected:
            raise RuntimeError('Tag lease rejected')
        self.writes.append(('tag', target))
        self.tag = target
        if target == NEW:
            self.fail('tag-response')

    def upload(self, directory):
        self.writes.append(('upload', directory.name))
        uploaded = release.package_hashes(directory)
        if directory == Path(self.args.package).resolve() and self.fail_at == 'partial-upload' and not self.fail_used:
            self.hashes[release.FILES[0]] = uploaded[release.FILES[0]]
            self.fail('partial-upload')
        self.hashes = uploaded

    def download(self, directory, expected):
        if directory.name == 'replacement-download':
            self.fail('draft-readback')
        if directory.name == 'published-download':
            self.fail('public-readback')
        if self.hashes != expected:
            raise ValueError('Download mismatch')


class ReplacementTests(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        self.addCleanup(self.temp.cleanup)
        self.root = Path(self.temp.name)
        candidate = self.root / 'candidate'
        old = self.root / 'old'
        candidate_hashes = make_package(candidate, b'candidate')
        old_hashes = make_package(old, b'old')
        self.args = types.SimpleNamespace(repository='example/project', source_sha=SOURCE,
            expected_main=SOURCE, expected_tag_object=OLD, expected_release_updated_at='before',
            expected_exe_sums_sha256=candidate_hashes['SHA256SUMS.txt'],
            package=str(candidate), backup=str(self.root / 'backup'))
        self.remote = Remote(self.args, old, old_hashes)
        self.subject = release.Release(self.args)
        for name in ('api', 'push_tag', 'upload', 'download'):
            replacement = getattr(self.remote, name)
            self.addCleanup(patch.stopall)
            patch.object(self.subject, name, replacement).start()
        self.backup = self.subject.backup
        self.backup.mkdir()
        shutil.copytree(old, self.backup / 'assets')
        previous = self.remote.snapshot()
        release.save(self.backup / 'remote-before.json', previous)
        release.save(self.backup / 'prepared.json', {'source': SOURCE, 'repository': self.args.repository,
            'package': candidate_hashes, 'previous': old_hashes, 'notes': 'New reviewed notes',
            'state': release.Release.fingerprint(previous)})
        self.original = copy.deepcopy(self.remote.metadata)
        self.original_hashes = dict(old_hashes)
        def git_result(*args, **kwargs):
            if args[1] == 'mktag':
                return NEW
            if args[1:3] == ('bundle', 'list-heads'):
                return f'{OLD} refs/tags/{release.TAG}'
            return SOURCE
        patch.object(release, 'run', side_effect=git_result).start()

    def test_draft_state_uses_archived_release_identity(self):
        self.remote.metadata['draft'] = True
        with self.assertRaisesRegex(RuntimeError, 'HTTP 404'):
            self.remote.api(f'releases/tags/{release.TAG}')
        state = self.subject.state()
        self.assertEqual(42, state['release']['id'])
        self.assertTrue(state['release']['draft'])
        self.assertIn('untagged-', state['release']['html_url'])
        self.assertEqual(OLD, state['tag']['sha'])

    def test_initial_state_finds_public_release_before_backup_exists(self):
        args = copy.copy(self.args)
        args.backup = str(self.root / 'not-yet-prepared')
        subject = release.Release(args)
        with patch.object(subject, 'api', self.remote.api):
            self.assertEqual(self.remote.snapshot(), subject.state())

    def test_success_preserves_identity_and_publishes_verified_package(self):
        self.subject.apply()
        self.assertEqual(NEW, self.remote.tag)
        self.assertFalse(self.remote.metadata['draft'])
        self.assertEqual(42, self.remote.metadata['id'])
        self.assertEqual(release.package_hashes(self.subject.package), self.remote.hashes)
        self.assertIn(SOURCE, self.remote.metadata['target_commitish'])
        self.assertEqual('VERIFIED', json.loads((self.backup / 'replacement.json').read_text())['status'])

    def test_each_partial_failure_restores_original_release(self):
        for point in ('draft-response', 'tag-response', 'partial-upload', 'draft-readback', 'publish-response', 'public-readback'):
            with self.subTest(point=point):
                self.remote.metadata = copy.deepcopy(self.original)
                self.remote.tag = OLD
                self.remote.hashes = dict(self.original_hashes)
                self.remote.fail_at, self.remote.fail_used = point, False
                with self.assertRaisesRegex(RuntimeError, 'previous release was restored'):
                    self.subject.apply()
                self.assertTrue(self.remote.fail_used)
                self.assertEqual(self.original, self.remote.metadata)
                self.assertEqual(OLD, self.remote.tag)
                self.assertEqual(self.original_hashes, self.remote.hashes)

    def test_changed_remote_state_stops_before_any_write(self):
        for field, value in (('main', 'd' * 40), ('tag', 'd' * 40), ('body', 'Changed elsewhere')):
            with self.subTest(field=field):
                self.remote.main, self.remote.tag = SOURCE, OLD
                self.remote.metadata = copy.deepcopy(self.original)
                if field in ('main', 'tag'):
                    setattr(self.remote, field, value)
                else:
                    self.remote.metadata[field] = value
                with self.assertRaises(ValueError):
                    self.subject.apply()
                self.assertEqual([], self.remote.writes)

    def test_failure_before_upload_preserves_existing_asset_ids(self):
        self.remote.fail_at = 'tag-response'
        with self.assertRaisesRegex(RuntimeError, 'previous release was restored'):
            self.subject.apply()
        self.assertFalse(any(kind == 'upload' for kind, _ in self.remote.writes))
        self.assertEqual(self.original_hashes, self.remote.hashes)

    def test_changed_candidate_or_backup_stops_before_any_write(self):
        for folder in (self.subject.package, self.backup / 'assets'):
            path = folder / release.FILES[0]
            before = path.read_bytes()
            path.write_bytes(b'corrupt')
            with self.assertRaises(ValueError):
                self.subject.apply()
            self.assertEqual([], self.remote.writes)
            path.write_bytes(before)

    def test_immutable_release_is_never_edited(self):
        self.remote.metadata['immutable'] = True
        with self.assertRaisesRegex(ValueError, 'mutable'):
            self.subject.apply()
        self.assertEqual([], self.remote.writes)

    def test_changed_archived_metadata_stops_before_any_write(self):
        path = self.backup / 'remote-before.json'
        previous = json.loads(path.read_text())
        previous['release']['body'] = 'Substituted archived notes'
        release.save(path, previous)
        with self.assertRaisesRegex(ValueError, 'metadata changed'):
            self.subject.apply()
        self.assertEqual([], self.remote.writes)

    def test_accepted_exe_manifest_cannot_be_substituted(self):
        self.args.expected_exe_sums_sha256 = '0' * 64
        with self.assertRaisesRegex(ValueError, 'accepted package'):
            self.subject.apply()
        self.assertEqual([], self.remote.writes)

    def test_concurrent_tag_change_stops_recovery(self):
        self.remote.tag = 'd' * 40
        with self.assertRaisesRegex(RuntimeError, 'Concurrent remote change'):
            self.subject.restore(json.loads((self.backup / 'remote-before.json').read_text()), NEW)
        self.assertEqual([], self.remote.writes)

    def test_recovery_failure_is_reported_and_draft_is_not_published(self):
        self.remote.fail_at = 'partial-upload'
        original_upload = self.remote.upload
        def upload(directory):
            if directory == self.backup / 'assets':
                raise RuntimeError('Recovery upload unavailable')
            return original_upload(directory)
        with patch.object(self.subject, 'upload', upload):
            with self.assertRaisesRegex(RuntimeError, 'recovery failed'):
                self.subject.apply()
        self.assertTrue(self.remote.metadata['draft'])
        self.assertTrue((self.backup / 'rollback-error.json').exists())

    def test_manual_recovery_after_interruption_uses_archived_files(self):
        self.remote.tag = NEW
        self.remote.metadata['draft'] = True
        self.remote.hashes = release.package_hashes(self.subject.package)
        release.save(self.backup / 'replacement-tag.json', {'sha': NEW, 'source': SOURCE})
        self.subject.recover()
        self.assertEqual(OLD, self.remote.tag)
        self.assertEqual(self.original_hashes, self.remote.hashes)
        self.assertEqual(self.original, self.remote.metadata)

    def test_source_zip_is_checked_by_git_blob_content(self):
        path = self.subject.package / release.FILES[2]
        payload = b'reviewed source\r\n'
        oid = hashlib.sha1(b'blob ' + str(len(payload)).encode() + b'\0' + payload).hexdigest()
        with zipfile.ZipFile(path, 'w') as archive:
            archive.comment = SOURCE.encode()
            archive.writestr('source.cs', payload)
        with patch.object(release, 'run', return_value=f'100644 blob {oid}\tsource.cs\0'):
            release.verify_sources(self.subject.package, SOURCE)
            with zipfile.ZipFile(path, 'w') as archive:
                archive.comment = SOURCE.encode()
                archive.writestr('source.cs', b'wrong source')
            with self.assertRaisesRegex(ValueError, 'content differs'):
                release.verify_sources(self.subject.package, SOURCE)


class LocalGitTests(unittest.TestCase):
    def test_release_body_links_point_to_the_selected_source_revision(self):
        text = '[Docs](../README.md) [Русский](./v0.0.4.ru.md) [Changes](../../CHANGELOG.md) [Section](#fixed) [Download](https://example.invalid/file.exe)'
        result = release.release_notes(text, 'example/project', SOURCE)
        for path in ('docs/README.md', 'docs/releases/v0.0.4.ru.md', 'CHANGELOG.md'):
            self.assertIn(f'https://github.com/example/project/blob/{SOURCE}/{path}', result)
        self.assertIn('[Section](#fixed)', result)
        self.assertIn('[Download](https://example.invalid/file.exe)', result)

    def test_native_git_accepts_tag_bytes_and_rejects_stale_lease(self):
        # Real Git, entirely local: Windows newline handling and tag-object leases.
        with tempfile.TemporaryDirectory() as folder:
            root = Path(folder).resolve()
            work, origin = root / 'work', root / 'origin.git'
            release.run('git', 'init', str(work))
            release.run('git', 'init', '--bare', str(origin))
            git = ('git', '-C', str(work))
            release.run(*git, 'config', 'user.name', 'Release regression')
            release.run(*git, 'config', 'user.email', 'release@example.invalid')
            release.run(*git, 'remote', 'add', 'origin', str(origin))
            tree = release.run(*git, 'mktree', data='')
            commit = release.run(*git, 'commit-tree', tree, '-m', 'Isolated release fixture')
            def tag(message):
                return release.run(*git, 'mktag', data=f'object {commit}\ntype commit\ntag v0.0.4\ntagger Release regression <release@example.invalid> 1700000000 +0000\n\n{message}\n')
            old_tag, new_tag = tag('original'), tag('replacement')
            release.run(*git, 'push', 'origin', f'{old_tag}:refs/tags/v0.0.4')
            release.run(*git, 'push', 'origin', f'{new_tag}:refs/tags/v0.0.4', f'--force-with-lease=refs/tags/v0.0.4:{old_tag}')
            with self.assertRaises(RuntimeError):
                release.run(*git, 'push', 'origin', f'{old_tag}:refs/tags/v0.0.4', f'--force-with-lease=refs/tags/v0.0.4:{old_tag}')
            actual = release.run('git', '--git-dir', str(origin), 'rev-parse', 'refs/tags/v0.0.4')
            self.assertEqual(new_tag, actual)


if __name__ == '__main__':
    unittest.main()
