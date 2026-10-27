using System.Net;
using System.Text.Json;

namespace DeviceTweakerQa;

internal static class WorkbenchHtml
{
    private const string Header = "<!doctype html><meta charset=utf-8><title>DEVICE TWEAKER QA</title><style>body{background:#101014;color:#ddd;font:15px Consolas,monospace;margin:24px}img{max-width:100%;border:1px solid #555}section{margin:22px 0;padding:16px;border:1px solid #444}button,input{margin:8px}pre{white-space:pre-wrap}.row{display:flex;gap:8px}.row>div{width:33%}a{color:#ace}</style>";
    public static void Session(string output, DevRun run)
    {
        string text = Header + "<h1>" + WebUtility.HtmlEncode(run.ScenarioId + " / " + run.Status) + "</h1><p>Read-only sandbox. Native client captures; synthetic hardware. Physical input and live driver behavior are not tested.</p>";
        foreach (DevCapture c in run.Captures.Where(c => !c.Timeline))
            text += $"<section><h2>{WebUtility.HtmlEncode(c.Id)}</h2><a href='{WebUtility.HtmlEncode(c.Tree)}'>Control tree / state</a><p><img src='{WebUtility.HtmlEncode(c.Path)}'></section>";
        var frames = run.Captures.Where(c => c.Timeline).Select(c => new { path = c.Path, time = c.TimeMs, id = c.Id }).ToArray();
        if (frames.Length != 0)
        {
            text += "<section><h2>Transition frames</h2><button id=play>Play / pause</button><input id=slider type=range min=0 value=0 max='" + (frames.Length - 1) + "'><pre id=stamp></pre><img id=frame></section>";
            text += "<script>const frames=" + JsonSerializer.Serialize(frames) + ";const s=document.getElementById('slider'),im=document.getElementById('frame'),st=document.getElementById('stamp');let timer;function show(){const f=frames[+s.value];im.src=f.path;st.textContent=f.id+' / '+f.time.toFixed(1)+' ms';}s.oninput=()=>{clearTimeout(timer);timer=null;show()};function tick(){let n=+s.value;if(n>=frames.length-1){timer=null;return;}let delay=Math.max(20,frames[n+1].time-frames[n].time);timer=setTimeout(()=>{s.value=n+1;show();tick()},delay)}document.getElementById('play').onclick=()=>{if(timer){clearTimeout(timer);timer=null;}else{if(+s.value>=frames.length-1)s.value=0;show();tick()}};show();</script>";
        }
        text += "<h2>Steps / exceptions</h2><pre>" + WebUtility.HtmlEncode(JsonSerializer.Serialize(run, DevFiles.Json)) + "</pre>";
        if (Directory.Exists(Path.Combine(output, "timelines")))
            foreach (string contact in Directory.EnumerateFiles(Path.Combine(output, "timelines"), "*-contact-*.png"))
                text += "<section><h2>Frame contact sheet</h2><img src='timelines/" + WebUtility.HtmlEncode(Path.GetFileName(contact)) + "'></section>";
        File.WriteAllText(Path.Combine(output, "index.html"), text);
    }
    public static void Comparison(string output, object result)
    {
        using JsonDocument doc = JsonDocument.Parse(JsonSerializer.Serialize(result, DevFiles.Json));
        string text = Header + "<h1>" + doc.RootElement.GetProperty("status").GetString() + "</h1><p>Left: before. Middle: after. Right: changed pixels. Changes require inspection; no baseline is replaced.</p>";
        foreach (JsonElement entry in doc.RootElement.GetProperty("checkpoints").EnumerateArray())
        {
            text += "<section><h2>" + WebUtility.HtmlEncode(entry.GetProperty("id").GetString()) + "</h2><pre>" + WebUtility.HtmlEncode(entry.GetProperty("difference").ToString()) + "</pre><div class=row>";
            foreach (string key in new[] { "before", "after", "diff" })
            { string? path = entry.GetProperty(key).GetString(); text += path is null ? "<div>Size changed</div>" : "<div><img src='" + WebUtility.HtmlEncode(path) + "'></div>"; }
            text += "</div></section>";
        }
        File.WriteAllText(Path.Combine(output, "index.html"), text);
    }
}
