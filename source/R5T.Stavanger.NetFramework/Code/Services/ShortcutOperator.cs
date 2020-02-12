using System;

using IWshRuntimeLibrary; // For embedded interop types to be true, use the WshShell interface type instead of the WshShellClass type.


namespace R5T.Stavanger.NetFramework
{
    public class ShortcutOperator : IShortcutOperator
    {
        private IShortcutPathConventions ShortcutPathConventions { get; }


        public ShortcutOperator(IShortcutPathConventions shortcutPathConventions)
        {
            this.ShortcutPathConventions = shortcutPathConventions;
        }

        private void ModifyShortcut(string shortcutFilePath, Action<IWshShortcut> shortcutAction)
        {
            WshShell windowsShell = new WshShell();
            IWshShortcut shortcut = windowsShell.CreateShortcut(shortcutFilePath) as IWshShortcut;

            shortcutAction(shortcut);

            shortcut.Save();
        }

        private T FunctionOnShortcut<T>(string shortcutFilePath, Func<IWshShortcut, T> shortcutFunction)
        {
            WshShell windowsShell = new WshShell();
            IWshShortcut shortcut = windowsShell.CreateShortcut(shortcutFilePath) as IWshShortcut;

            var output = shortcutFunction(shortcut);

            return output;
        }

        public string CreateShortcut(string shortcutFilePath, string shortcutTargetPath)
        {
            string shortcutLinkFilePath = this.ShortcutPathConventions.MakeFilePathIntoLinkFilePath(shortcutFilePath);

            this.ModifyShortcut(shortcutFilePath, (shortcut) =>
            {
                shortcut.TargetPath = shortcutTargetPath;
            });

            return shortcutLinkFilePath;
        }

        public string CreateShortcut(string shortcutFilePath, string shortcutTargetPath, string description)
        {
            string shortcutLinkFilePath = this.ShortcutPathConventions.MakeFilePathIntoLinkFilePath(shortcutFilePath);

            this.ModifyShortcut(shortcutFilePath, (shortcut) =>
            {
                shortcut.TargetPath = shortcutTargetPath;
                shortcut.Description = description;
            });

            return shortcutLinkFilePath;
        }

        public string GetShortcutTargetPath(string shortcutFilePath)
        {
            var shortcutTargetPath = this.FunctionOnShortcut(shortcutFilePath, (shortcut) =>
            {
                var output = shortcut.TargetPath;
                return output;
            });

            return shortcutTargetPath;
        }

        public void SetShortcutTargetPath(string shortcutFilePath, string shortcutTargetPath)
        {
            this.ModifyShortcut(shortcutFilePath, (shortcut) =>
            {
                shortcut.TargetPath = shortcutTargetPath;
            });
        }

        public string GetShortcutDescription(string shortcutFilePath)
        {
            var shortcutDescription = this.FunctionOnShortcut(shortcutFilePath, (shortcut) =>
            {
                var output = shortcut.Description;
                return output;
            });

            return shortcutDescription;
        }

        public void SetShortcutDescription(string shortcutFilePath, string description)
        {
            this.ModifyShortcut(shortcutFilePath, (shortcut) =>
            {
                shortcut.Description = description;
            });
        }
    }
}
