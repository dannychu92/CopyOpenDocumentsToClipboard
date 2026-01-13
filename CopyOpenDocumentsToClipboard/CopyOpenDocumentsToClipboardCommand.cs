using Microsoft.VisualStudio.Shell;
using System;
using System.ComponentModel.Design;
using System.Text;
using System.Threading.Tasks;

namespace CopyOpenDocumentsToClipboard
{
    internal sealed class CopyOpenDocumentsToClipboardCommand
    {
        public const int CommandId_CopyAllOpenDocuments = 0x0100;
        public const int CommandId_CopySingleDocument = 0x0101;

        public static readonly Guid CommandSet = new Guid("c0bdb4d1-17b6-4a33-9f06-2d73f6d3c3a7");

        private readonly AsyncPackage _package;

        private CopyOpenDocumentsToClipboardCommand(AsyncPackage package, OleMenuCommandService commandService)
        {
            _package = package;

            // Tools -> Copy open documents to Clipboard
            var menuCommandIdCopyAll = new CommandID(CommandSet, CommandId_CopyAllOpenDocuments);
            var menuItemCopyAll = new OleMenuCommand(ExecuteCopyAllOpenDocuments, menuCommandIdCopyAll);
            commandService.AddCommand(menuItemCopyAll);

            // Document tab right-click -> Copy document to Clipboard
            var menuCommandIdCopySingle = new CommandID(CommandSet, CommandId_CopySingleDocument);
            var menuItemCopySingle = new OleMenuCommand(ExecuteCopySingleDocument, menuCommandIdCopySingle);
            commandService.AddCommand(menuItemCopySingle);
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

        private void ExecuteCopyAllOpenDocuments(object sender, EventArgs e)
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            _package.JoinableTaskFactory.RunAsync(async () =>
            {
                await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync();

                var dte = await GetDteAsync();
                if (dte == null)
                    return;

                var solutionDir = TryGetSolutionDir(dte);

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

                    if (TryGetTextDocumentContent(doc, out var content) == false)
                        continue;

                    var block = BuildSingleDocumentBlock(dte, doc, content, solutionDir);
                    if (string.IsNullOrWhiteSpace(block))
                        continue;

                    if (addedAny == true)
                        sb.AppendLine();

                    sb.Append(block);
                    addedAny = true;
                }

                if (addedAny == false)
                    return;

                System.Windows.Forms.Clipboard.SetText(sb.ToString());
            }).FileAndForget("CopyOpenDocumentsToClipboard/ExecuteCopyAllOpenDocuments");
        }

        private void ExecuteCopySingleDocument(object sender, EventArgs e)
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            _package.JoinableTaskFactory.RunAsync(async () =>
            {
                await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync();

                var dte = await GetDteAsync();
                if (dte == null)
                    return;

                var doc = dte.ActiveDocument;
                if (doc == null)
                    return;

                if (TryGetTextDocumentContent(doc, out var content) == false)
                    return;

                var solutionDir = TryGetSolutionDir(dte);

                var block = BuildSingleDocumentBlock(dte, doc, content, solutionDir);
                if (string.IsNullOrWhiteSpace(block))
                    return;

                System.Windows.Forms.Clipboard.SetText(block);
            }).FileAndForget("CopyOpenDocumentsToClipboard/ExecuteCopySingleDocument");
        }



        private async Task<EnvDTE80.DTE2> GetDteAsync()
        {
            // DTE is a UI-thread COM service. Assert and ensure we are on UI thread.
            await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync();

            var dteObj = await _package.GetServiceAsync(typeof(EnvDTE.DTE));
            return dteObj as EnvDTE80.DTE2;
        }

        private static bool TryGetTextDocumentContent(EnvDTE.Document doc, out string content)
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            content = string.Empty;

            EnvDTE.TextDocument textDoc;

            try
            {
                textDoc = doc.Object("TextDocument") as EnvDTE.TextDocument;
            }
            catch
            {
                return false;
            }

            if (textDoc == null)
                return false;

            try
            {
                var start = textDoc.StartPoint.CreateEditPoint();
                var end = textDoc.EndPoint.CreateEditPoint();
                content = start.GetText(end);
                return true;
            }
            catch
            {
                return false;
            }
        }

        private static string BuildSingleDocumentBlock(EnvDTE80.DTE2 dte, EnvDTE.Document doc, string content, string solutionDir)
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            var displayPath = doc.FullName ?? doc.Name;

            // Prefer project-relative paths: "<ProjectName>/path/inside/project"
            if (string.IsNullOrWhiteSpace(doc.FullName) == false)
            {
                try
                {
                    EnvDTE.ProjectItem projectItem = null;

                    try
                    {
                        if (dte.Solution != null)
                        {
                            projectItem = dte.Solution.FindProjectItem(doc.FullName);
                        }
                    }
                    catch
                    {
                        projectItem = null;
                    }

                    var owningProject = projectItem != null ? projectItem.ContainingProject : null;

                    if (owningProject != null && string.IsNullOrWhiteSpace(owningProject.FullName) == false)
                    {
                        var projectDir = System.IO.Path.GetDirectoryName(owningProject.FullName) ?? "";
                        if (string.IsNullOrWhiteSpace(projectDir) == false)
                        {
                            string relativeToProject;

                            try
                            {
                                relativeToProject = MakeRelativePath(projectDir, doc.FullName);
                            }
                            catch
                            {
                                relativeToProject = doc.FullName;
                            }

                            relativeToProject = relativeToProject.Replace('\\', '/').TrimStart('/');

                            var projectName = owningProject.Name ?? "";
                            if (string.IsNullOrWhiteSpace(projectName) == false)
                            {
                                displayPath = projectName + "/" + relativeToProject;
                            }
                            else
                            {
                                displayPath = relativeToProject;
                            }
                        }
                    }
                    else if (string.IsNullOrWhiteSpace(solutionDir) == false)
                    {
                        // Fallback: solution-relative
                        try
                        {
                            var full = doc.FullName;
                            var dir = solutionDir.TrimEnd('\\') + "\\";
                            if (full.StartsWith(dir, StringComparison.OrdinalIgnoreCase))
                            {
                                displayPath = full.Substring(dir.Length).Replace('\\', '/');
                            }
                            else
                            {
                                displayPath = full.Replace('\\', '/');
                            }
                        }
                        catch
                        {
                            displayPath = doc.FullName ?? doc.Name;
                        }
                    }
                    else
                    {
                        displayPath = doc.FullName ?? doc.Name;
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

            var sb = new StringBuilder();

            sb.Append("// === FILE: ");
            sb.Append(fileName);

            if (string.IsNullOrWhiteSpace(pathDir) == false)
            {
                sb.Append(" | PATH: ");
                sb.Append(pathDir);
            }

            sb.AppendLine(" ================================");
            sb.AppendLine(content);

            return sb.ToString();
        }

        private static string TryGetSolutionDir(EnvDTE80.DTE2 dte)
        {
            ThreadHelper.ThrowIfNotOnUIThread();

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

            return solutionDir;
        }

        static string MakeRelativePath(string baseDir, string fullPath)
        {
            try
            {
                if (string.IsNullOrWhiteSpace(baseDir) || string.IsNullOrWhiteSpace(fullPath))
                    return fullPath;

                if (baseDir.EndsWith("\\") == false)
                    baseDir += "\\";

                var baseUri = new Uri(baseDir, UriKind.Absolute);
                var fileUri = new Uri(fullPath, UriKind.Absolute);

                if (baseUri.Scheme != fileUri.Scheme)
                    return fullPath;

                var relativeUri = baseUri.MakeRelativeUri(fileUri);
                var relativePath = Uri.UnescapeDataString(relativeUri.ToString());

                return relativePath.Replace('/', '\\');
            }
            catch
            {
                return fullPath;
            }
        }
    }
}
