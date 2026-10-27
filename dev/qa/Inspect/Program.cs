using System.Reflection;
using System.Security.Cryptography;
using System.Text.Json;
using Microsoft.CodeAnalysis;
using Microsoft.CodeAnalysis.CSharp;
using Microsoft.CodeAnalysis.CSharp.Syntax;

// Read source and call only the pure translation function. Never start MainForm.
Environment.SetEnvironmentVariable("DEVICE_TWEAKER_QA_READ_ONLY", "1");
Environment.SetEnvironmentVariable("DEVICE_TWEAKER_LANGUAGE", "RU");
if (args.FirstOrDefault() == "--dev-validate")
{
    var scene = DeviceTweakerQa.DevFiles.Read<DeviceTweakerQa.DevScenario>(args[1]); scene.Validate();
    Console.WriteLine(System.Text.Json.JsonSerializer.Serialize(scene, DeviceTweakerQa.DevFiles.Json)); return;
}
if (args.FirstOrDefault() == "--dev-compare")
{
    object result = DeviceTweakerQa.VisualComparison.CompareRuns(args[1], args[2], args[3]);
    Console.WriteLine(System.Text.Json.JsonSerializer.Serialize(result, DeviceTweakerQa.DevFiles.Json)); return;
}
if (args.FirstOrDefault() is "--dev-verify" or "--dev-verify-complete")
{
    var run = DeviceTweakerQa.DevFiles.Read<DeviceTweakerQa.DevRun>(Path.Combine(args[1], "session.json"));
    DeviceTweakerQa.DevFiles.Verify(args[1], run);
    if (args[0] == "--dev-verify-complete") DeviceTweakerQa.DevFiles.VerifyExecution(args[1]);
    Console.WriteLine("Session evidence verified"); return;
}
if (args.FirstOrDefault() == "--dev-review-template")
{ DeviceTweakerQa.DevelopmentReview.Template(args[1], args[2]); return; }
if (args.FirstOrDefault() == "--dev-review")
{
    string status = DeviceTweakerQa.DevelopmentReview.Complete(args[1], args[2]);
    Console.WriteLine(status); Environment.ExitCode = status == "REVIEWED" ? 0 : 1; return;
}
if (args.FirstOrDefault() == "--controls")
{
    int code = 1;
    Exception? failure = null;
    Thread worker = new(() => { try { code = ControlContracts.Run(args.Skip(1).ToArray()); } catch (Exception ex) { failure = ex; } });
    worker.SetApartmentState(ApartmentState.STA);
    worker.Start();
    worker.Join();
    if (failure is not null) Console.Error.WriteLine(failure);
    Environment.Exit(code);
    return;
}
if (args.Length != 3) throw new ArgumentException("workspace, production DLL, output JSON required");
string workspace = Path.GetFullPath(args[0]);
Assembly assembly = Assembly.LoadFrom(Path.GetFullPath(args[1]));
Type language = assembly.GetType("DeviceTweakerCS.UiLanguage", throwOnError: true)!;
MethodInfo translate = language.GetMethod("Text", [typeof(string)])!;
string Russian(string source) => (string)translate.Invoke(null, [source])!;
string[] folders = ["Affinity", "Core", "Devices", "GUI", "Interop", "Localization", "Models", "Tweaks"];
List<object> files = [], literals = [], candidates = [];
string[] properties = ["Text", "PlaceholderText", "AccessibleName", "AccessibleDescription"];
string[] calls = ["Text", "SetToolTip", "ShowThemedInfo", "ShowThemedError", "ShowThemedWarning", "ShowThemedConfirm", "ShowOperationResult", "SetProgress", "NewDialogButton", "NewTopButton", "AddError", "AddWarning", "AddSkipped", "AddSuccess", "SetDevicesBusyStage", "BeginDevicesBusy"];
string? ConstantString(ExpressionSyntax expression) => expression switch
{
    LiteralExpressionSyntax literal when literal.IsKind(SyntaxKind.StringLiteralExpression) => literal.Token.ValueText,
    BinaryExpressionSyntax binary when binary.IsKind(SyntaxKind.AddExpression)
        && ConstantString(binary.Left) is string left && ConstantString(binary.Right) is string right => left + right,
    ParenthesizedExpressionSyntax paren => ConstantString(paren.Expression),
    _ => null
};
foreach (string folder in folders)
foreach (string path in Directory.GetFiles(Path.Combine(workspace, "DEVICE TWEAKER", folder), "*.cs"))
{
    string source = File.ReadAllText(path);
    var syntax = CSharpSyntaxTree.ParseText(source, new CSharpParseOptions(LanguageVersion.Preview));
    SyntaxNode root = syntax.GetRoot();
    string relative = Path.GetRelativePath(workspace, path).Replace('\\', '/');
    var methods = root.DescendantNodes().OfType<BaseMethodDeclarationSyntax>().Select(m => new
    {
        name = m is MethodDeclarationSyntax method ? method.Identifier.ValueText : m.Kind().ToString(),
        line = m.GetLocation().GetLineSpan().StartLinePosition.Line + 1,
        endLine = m.GetLocation().GetLineSpan().EndLinePosition.Line + 1
    }).ToArray();
    files.Add(new { path = relative, lines = source.Split('\n').Length, methods,
        sha256 = Convert.ToHexString(SHA256.HashData(File.ReadAllBytes(path))) });
    foreach (var node in root.DescendantNodes().OfType<InterpolatedStringExpressionSyntax>())
    {
        var assignment = node.Ancestors().OfType<AssignmentExpressionSyntax>().FirstOrDefault();
        var invocation = node.Ancestors().OfType<InvocationExpressionSyntax>().FirstOrDefault(i => calls.Contains(i.Expression.ToString().Split('.').Last()));
        string? property = assignment?.Left.ToString().Split('.').Last();
        if (!properties.Contains(property) && invocation is null) continue;
        string template = string.Concat(node.Contents.Select(c => c is InterpolatedStringTextSyntax text ? text.TextToken.ValueText : "1"));
        string ru = Russian(template);
        candidates.Add(new { path = relative, line = node.GetLocation().GetLineSpan().StartLinePosition.Line + 1,
            value = template, complete = template, russian = ru, unchanged = ru == template,
            context = node.ToString(), interpolated = true });
    }
    foreach (var node in root.DescendantNodes().OfType<LiteralExpressionSyntax>().Where(n => n.IsKind(SyntaxKind.StringLiteralExpression)))
    {
        string value = node.Token.ValueText;
        int line = node.GetLocation().GetLineSpan().StartLinePosition.Line + 1;
        var assignment = node.Ancestors().OfType<AssignmentExpressionSyntax>().FirstOrDefault();
        var invocation = node.Ancestors().OfType<InvocationExpressionSyntax>().FirstOrDefault(i => calls.Contains(i.Expression.ToString().Split('.').Last()));
        string? property = assignment?.Left.ToString().Split('.').Last();
        string? call = invocation?.Expression.ToString().Split('.').Last();
        bool candidate = properties.Contains(property) || calls.Contains(call);
        var argument = node.Ancestors().OfType<ArgumentSyntax>().FirstOrDefault(a => a.Parent?.Parent == invocation);
        ExpressionSyntax expression = argument?.Expression ?? assignment?.Right ?? node;
        string complete = ConstantString(expression) ?? value;
        string ru = candidate ? Russian(complete) : "";
        var entry = new { path = relative, line, value, property, call, candidate, russian = ru,
            complete, unchanged = candidate && complete == ru, context = expression.ToString() };
        literals.Add(entry);
        if (candidate && value.Length > 0) candidates.Add(entry);
    }
}
File.WriteAllText(Path.GetFullPath(args[2]), JsonSerializer.Serialize(new
{
    scope = "Syntax inventory of active production C#; literal candidates plus pure RU translation. Interpolated and indirect UI strings require separate review.",
    assemblyPath = assembly.Location,
    assemblySha256 = Convert.ToHexString(SHA256.HashData(File.ReadAllBytes(assembly.Location))),
    files, literals, candidates
}, new JsonSerializerOptions { WriteIndented = true }));
Console.WriteLine($"Source files: {files.Count}; string literals: {literals.Count}; UI literal candidates: {candidates.Count}");
