using Microsoft.VisualStudio.Shell;
using System;
using System.Collections.Generic;
using System.ComponentModel.Design;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CopyOpenDocumentsToClipboard
{
    internal sealed class CopyOpenDocumentsToClipboardCommand
    {
        public const int CommandId = 0x0100;
        public static readonly Guid CommandSet = new Guid("c0bdb4d1-17b6-4a33-9f06-2d73f6d3c3a7");

        private readonly AsyncPackage _package;

        private CopyOpenDocumentsToClipboardCommand(AsyncPackage package, OleMenuCommandService commandService)
        {
            _package = package;

            var menuCommandId = new CommandID(CommandSet, CommandId);
            var menuItem = new OleMenuCommand(Execute, menuCommandId);
            commandService.AddCommand(menuItem);
        }

        public static CopyOpenDocumentsToClipboardCommand Instance { get; private set; }

        public static async Task InitializeAsync(AsyncPackage package)
        {
            await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync();

            var commandService = await package.GetServiceAsync(typeof(IMenuCommandService)) as OleMenuCommandService;
            if (commandService == null)
                return;

            Instance = new CopyOpenDocumentsToClipboardCommand(package, commandService);
        }

        private void Execute(object sender, EventArgs e)
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            ThreadHelper.JoinableTaskFactory.RunAsync(async () =>
            {
                await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync();

                var dteObj = await _package.GetServiceAsync(typeof(EnvDTE.DTE));
                var dte = dteObj as EnvDTE80.DTE2;
                if (dte == null)
                    return;

                var solutionDir = "";
                try
                {
                    if (dte.Solution != null && string.IsNullOrWhiteSpace(dte.Solution.FullName) == false)
                    {
                        solutionDir = System.IO.Path.GetDirectoryName(dte.Solution.FullName) ?? "";
                    }
                }
                catch
                {
                    solutionDir = "";
                }

                var sb = new StringBuilder();
                var addedAny = false;

                for (int i = 1; i <= dte.Documents.Count; i++)
                {
                    EnvDTE.Document doc;

                    try
                    {
                        doc = dte.Documents.Item(i);
                    }
                    catch
                    {
                        continue;
                    }

                    if (doc == null)
                        continue;

                    EnvDTE.TextDocument textDoc;

                    try
                    {
                        textDoc = doc.Object("TextDocument") as EnvDTE.TextDocument;
                    }
                    catch
                    {
                        continue;
                    }

                    if (textDoc == null)
                        continue;

                    var start = textDoc.StartPoint.CreateEditPoint();
                    var end = textDoc.EndPoint.CreateEditPoint();
                    var content = start.GetText(end);

                    var displayPath = doc.FullName ?? doc.Name;

                    if (string.IsNullOrWhiteSpace(solutionDir) == false && string.IsNullOrWhiteSpace(doc.FullName) == false)
                    {
                        try
                        {
                            var full = doc.FullName;
                            var dir = solutionDir.TrimEnd('\\') + "\\";
                            if (full.StartsWith(dir, StringComparison.OrdinalIgnoreCase))
                            {
                                displayPath = full.Substring(dir.Length);
                            }
                        }
                        catch
                        {
                            displayPath = doc.FullName ?? doc.Name;
                        }
                    }

                    var fileName = System.IO.Path.GetFileName(displayPath);
                    var pathDir = System.IO.Path.GetDirectoryName(displayPath);

                    if (string.IsNullOrWhiteSpace(pathDir) == false)
                    {
                        pathDir = pathDir.Replace('\\', '/');
                    }

                    if (addedAny == true)
                        sb.AppendLine();

                    sb.Append("// === FILE: ");
                    sb.Append(fileName);

                    if (string.IsNullOrWhiteSpace(pathDir) == false)
                    {
                        sb.Append(" | PATH: ");
                        sb.Append(pathDir);
                    }

                    sb.AppendLine(" ================================");
                    sb.AppendLine(content);

                    addedAny = true;
                }

                if (addedAny == false)
                    return;

                System.Windows.Forms.Clipboard.SetText(sb.ToString());
            });
        }

    }
}
