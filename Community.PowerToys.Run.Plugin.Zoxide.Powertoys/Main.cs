using ManagedCommon;

using System;
using System.Collections.Generic;
using System.IO;
using System.Windows;
using System.Windows.Input;

using Wox.Plugin;

using SimpleExec;

using System.Linq;

using Wox.Plugin.Logger;

using System.Diagnostics;

namespace Community.PowerToys.Run.Plugin.Zoxide.Powertoys;

/// <summary>
/// Main class of this plugin that implement all used interfaces.
/// </summary>
public class Main : IPlugin, IContextMenu, IDisposable
{
    /// <summary>
    /// ID of the plugin.
    /// </summary>
    public static string PluginID => "477BD4818E6134BCB38BB543A7B1E2EE";

    /// <summary>
    /// Name of the plugin.
    /// </summary>
    public string Name => "Zoxide";

    /// <summary>
    /// Description of the plugin.
    /// </summary>
    public string Description => "Bring zoxide to Powertoys";

    private PluginInitContext Context { get; set; }

    private string IconPath { get; set; }

    private bool Disposed { get; set; }

    private string ZoxidePath { get; set; }

    /// <summary>
    /// Return a filtered list, based on the given query.
    /// </summary>
    /// <param name="query">The query to filter the list.</param>
    /// <returns>A filtered list, can be empty when nothing was found.</returns>
    public List<Result> Query(Query query)
    {
        var search = query.Search;

        if (ZoxidePath is null)
        {
            return
            [
                new Result
                {
                    QueryTextDisplay = search,
                    IcoPath = IconPath,
                    Title = "Error: Cannot find zoxide",
                    SubTitle = "Error: Cannot find zoxide",
                    ToolTipData = new ToolTipData("Error", "Cannot find zoxide"),
                    Action = _ => true,
                    ContextData = search,
                }
            ];
        }

        string output, error;

        try
        {
            (output, error) = Command.ReadAsync(ZoxidePath, ["query", search]).Result;
        }
        catch (AggregateException e)
        {
            return
            [
                new Result
                {
                    QueryTextDisplay = search,
                    IcoPath = IconPath,
                    Title = "Error while query",
                    SubTitle = e.InnerExceptions.First() is ExitCodeReadException ex ? ex.StandardError.Trim() : e.Message.Trim(),
                    Action = _ =>
                    {
                        Clipboard.SetDataObject(e.InnerExceptions.First() is ExitCodeReadException ex2 ? ex2.StandardError.Trim() : e.Message.Trim());
                        return true;
                    },
                }
            ];
        }

        if (!(string.IsNullOrEmpty(error) || string.IsNullOrWhiteSpace(error)))
        {
            return
            [
                new Result
                {
                    QueryTextDisplay = search,
                    IcoPath = IconPath,
                    Title = "Error: zoxide",
                    SubTitle = error,
                    Action = _ =>
                    {
                        Clipboard.SetDataObject(error);
                        return true;
                    },
                }
            ];
        }

        var dir = new DirectoryInfo(output);
        return
        [
            new Result
            {
                QueryTextDisplay = search,
                IcoPath = IconPath,
                Title = "Navigate to: " + dir.Name.Trim(),
                SubTitle = "Full path: " + dir.FullName.Trim(),
                Action = _ =>
                {
                    Process.Start("explorer.exe", dir.FullName);
                    return true;
                }
            }
        ];
    }

    /// <summary>
    /// Initialize the plugin with the given <see cref="PluginInitContext"/>.
    /// </summary>
    /// <param name="context">The <see cref="PluginInitContext"/> for this plugin.</param>
    public void Init(PluginInitContext context)
    {
        Context = context ?? throw new ArgumentNullException(nameof(context));
        Context.API.ThemeChanged += OnThemeChanged;
        UpdateIconPath(Context.API.GetCurrentTheme());
        var pathEnv = Environment.GetEnvironmentVariable("PATH");
        if (string.IsNullOrEmpty(pathEnv))
        {
            return;
        }

        const string targetFile = "zoxide.exe";
        var zoxidePath = pathEnv
            .Split(Path.PathSeparator)
            .Where(p => !string.IsNullOrWhiteSpace(p))
            .Select(p => Path.Combine(p, targetFile))
            .FirstOrDefault(File.Exists);

        ZoxidePath = zoxidePath;
        Log.Info($"Current zoxide path: {ZoxidePath}", GetType());
    }

    /// <summary>
    /// Return a list context menu entries for a given <see cref="Result"/> (shown on the right side of the result).
    /// </summary>
    /// <param name="selectedResult">The <see cref="Result"/> for the list with context menu entries.</param>
    /// <returns>A list context menu entries.</returns>
    public List<ContextMenuResult> LoadContextMenus(Result selectedResult)
    {
        return [];
    }

    /// <inheritdoc/>
    public void Dispose()
    {
        Dispose(true);
        GC.SuppressFinalize(this);
    }

    /// <summary>
    /// Wrapper method for <see cref="Dispose()"/> that dispose additional objects and events form the plugin itself.
    /// </summary>
    /// <param name="disposing">Indicate that the plugin is disposed.</param>
    protected virtual void Dispose(bool disposing)
    {
        if (Disposed || !disposing)
        {
            return;
        }

        if (Context?.API != null)
        {
            Context.API.ThemeChanged -= OnThemeChanged;
        }

        Disposed = true;
    }

    private void UpdateIconPath(Theme theme) => IconPath = theme is Theme.Light or Theme.HighContrastWhite
        ? "Images/zoxide.powertoys.light.png"
        : "Images/zoxide.powertoys.dark.png";

    private void OnThemeChanged(Theme currentTheme, Theme newTheme) => UpdateIconPath(newTheme);
}