namespace SwapFileNamesShellExtension
{
    using System;
    using System.Collections.Generic;
    using System.Diagnostics.CodeAnalysis;
    using System.IO;
    using System.Linq;
    using System.Runtime.InteropServices;
    using System.Windows.Forms;

    using SharpShell.Attributes;
    using SharpShell.SharpContextMenu;

    /// <summary>
    ///     The context menu extension.
    /// </summary>
    [ComVisible(true)]
    [COMServerAssociation(AssociationType.AllFiles)]
    [SuppressMessage("Microsoft.Interoperability", "CA1405:ComVisibleTypeBaseTypesShouldBeComVisible")]
    public class SwapFileNamesContextMenu : SharpContextMenu
    {
        private static string capturedFileName;

        private bool isInDirectMode;

        /// <summary>
        ///     Determines whether this instance can show a shell context menu, given the specified selected file list.
        /// </summary>
        /// <returns>
        ///     <c>true</c> if this instance should show a shell context menu for the specified file list; otherwise, <c>false</c>.
        /// </returns>
        protected override bool CanShowMenu()
        {
            try
            {
                List<string> paths = this.SelectedItemPaths.ToList();
                if (paths.Count == 1 && File.Exists(paths[0]))
                {
                    isInDirectMode = false;
                    return true;
                }
                
                if (paths.Count == 2 && File.Exists(paths[0]) && File.Exists(paths[1]))
                {
                    isInDirectMode = true;
                    return true;
                }

                return false;
            }
            catch (Exception)
            {
                return false;
            }
        }

        /// <summary>
        ///     Creates the context menu. This can be a single menu item or a tree of them.
        /// </summary>
        /// <returns>
        ///     The context menu for the shell context menu.
        /// </returns>
        protected override ContextMenuStrip CreateMenu()
        {
            ContextMenuStrip menu = new ContextMenuStrip();

            if (isInDirectMode)
            {
                string swapFileNamesLabel = string.Format(Resources.SwapFileNamesDirectlyContextMenuItem, Path.GetFileName(capturedFileName));
                ToolStripMenuItem swapFileNamesMenuItem = new ToolStripMenuItem { Text = swapFileNamesLabel, Image = Resources.Icon };
                swapFileNamesMenuItem.Click += this.OnSwapFileNamesDirectlyClick;
                menu.Items.Add(swapFileNamesMenuItem);
            }
            else
            {
                if (capturedFileName != null && !capturedFileName.Equals(this.SelectedItemPaths.Single(), StringComparison.OrdinalIgnoreCase))
                {
                    string swapFileNamesLabel = string.Format(Resources.SwapFileNameContextMenuItem, Path.GetFileName(capturedFileName));
                    ToolStripMenuItem swapFileNamesMenuItem = new ToolStripMenuItem { Text = swapFileNamesLabel, Image = Resources.Icon };
                    swapFileNamesMenuItem.Click += this.OnSwapFileNamesClick;
                    menu.Items.Add(swapFileNamesMenuItem);
                }

                ToolStripMenuItem markFileNameMenuItem = new ToolStripMenuItem { Text = Resources.SelectFileNameContextMenuItem, Image = Resources.Icon };
                markFileNameMenuItem.Click += this.OnCaptureFileNameClick;
                menu.Items.Add(markFileNameMenuItem);
            }

            return menu;
        }

        private void OnSwapFileNamesDirectlyClick(object sender, EventArgs e)
        {
            try
            {
                capturedFileName = null;
                this.TrySwapFileNames(this.SelectedItemPaths.First(), this.SelectedItemPaths.Last());
            }
            catch (Exception exception)
            {
                MessageBox.Show(
                    string.Format(exception.Message),
                    Resources.ErrorMessageBoxTitle,
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
            }
        }

        private void OnCaptureFileNameClick(object sender, EventArgs eventArgs)
        {
            try
            {
                capturedFileName = this.SelectedItemPaths.Single();
            }
            catch (Exception exception)
            {
                MessageBox.Show(
                    string.Format(exception.Message),
                    Resources.ErrorMessageBoxTitle,
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
            }
        }

        private void OnSwapFileNamesClick(object sender, EventArgs e)
        {
            try
            {
                if (this.TrySwapFileNames(capturedFileName, this.SelectedItemPaths.Single()))
                {
                    capturedFileName = null;
                }
            }
            catch (Exception exception)
            {
                MessageBox.Show(
                    string.Format(exception.Message),
                    Resources.ErrorMessageBoxTitle,
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
            }
        }

        private bool TrySwapFileNames(string path1, string path2)
        {
            string directoryName1 = Path.GetDirectoryName(path1) ?? string.Empty;
            string directoryName2 = Path.GetDirectoryName(path2) ?? string.Empty;

            string fileName1 = Path.GetFileName(path1) ?? string.Empty;
            string fileName2 = Path.GetFileName(path2) ?? string.Empty;

            string tempPath1 = Path.Combine(directoryName1, Guid.NewGuid().ToString());
            string tempPath2 = Path.Combine(directoryName2, Guid.NewGuid().ToString());

            string destinationPath1 = Path.Combine(directoryName1, fileName2);
            string destinationPath2 = Path.Combine(directoryName2, fileName1);

            if (!directoryName1.Equals(directoryName2, StringComparison.OrdinalIgnoreCase)
                && (!CanCreateFile(destinationPath1) || !CanCreateFile(destinationPath2)))
            {
                return false;
            }

            File.Move(path1, tempPath1);
            File.Move(path2, tempPath2);
            File.Move(tempPath1, destinationPath1);
            File.Move(tempPath2, destinationPath2);

            return true;
        }

        private static bool CanCreateFile(string path)
        {
            if (!File.Exists(path))
            {
                return true;
            }

            DialogResult result = MessageBox.Show(
                string.Format(Resources.OverwriteExistingFileMessageBoxText, path),
                Resources.WarningMessageBoxTitle,
                MessageBoxButtons.YesNo,
                MessageBoxIcon.Warning,
                MessageBoxDefaultButton.Button2);

            if (result == DialogResult.Yes)
            {
                File.Delete(path);
                return true;
            }

            return false;
        }
    }
}
