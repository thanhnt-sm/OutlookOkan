// Stub file to provide base class inheritance for ThisAddIn when Designer.cs is excluded
// This allows compiling ThisAddIn.cs which uses 'override' on base class methods
namespace OutlookOkan
{
    public partial class ThisAddIn : Microsoft.Office.Tools.Outlook.AddInBase
    {
    }
}
