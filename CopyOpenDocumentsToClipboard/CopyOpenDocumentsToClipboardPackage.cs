using Microsoft.VisualStudio.Shell;
using System;
using System.Runtime.InteropServices;
using System.Threading;
using Task = System.Threading.Tasks.Task;

namespace CopyOpenDocumentsToClipboard
{
    [PackageRegistration(UseManagedResourcesOnly = true, AllowsBackgroundLoading = true)]
    [ProvideMenuResource("Commands.ctmenu", 1)]
    [Guid(CopyOpenDocumentsToClipboardPackage.PackageGuidString)]
    public sealed class CopyOpenDocumentsToClipboardPackage : AsyncPackage
    {
        public const string PackageGuidString = "8716c4fc-4f25-4908-83ff-b00ef81d3422";

        protected override async Task InitializeAsync(CancellationToken cancellationToken, IProgress<ServiceProgressData> progress)
        {
            await this.JoinableTaskFactory.SwitchToMainThreadAsync(cancellationToken);

            await CopyOpenDocumentsToClipboardCommand.InitializeAsync(this);
        }
    }
}
