namespace MonteCarlo.UI.Services;

/// <summary>
/// Light bridge between the Excel host layer and the UI settings view.
/// The UI project stays Excel-agnostic while the add-in project supplies
/// workbook persistence callbacks at runtime.
/// </summary>
public static class WorkbookSettingsBridge
{
    public static Func<WorkbookUserSettingsOverrides?>? LoadOverrides { get; set; }

    public static Action<WorkbookUserSettingsOverrides?>? SaveOverrides { get; set; }

    public static bool IsAvailable => LoadOverrides != null && SaveOverrides != null;

    public static WorkbookUserSettingsOverrides? Load() => LoadOverrides?.Invoke();

    public static void Save(WorkbookUserSettingsOverrides? overrides) =>
        SaveOverrides?.Invoke(overrides);
}
