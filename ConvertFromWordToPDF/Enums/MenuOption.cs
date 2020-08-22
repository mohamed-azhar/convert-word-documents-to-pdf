using System.ComponentModel.DataAnnotations;

namespace ConvertFromWordToPDF.Enums
{
    public enum MenuOption
    {
        [Display(Name = "Single Convert")]
        SingleConvert = 1,

        [Display(Name = "Multiple Convert")]
        MultipleConvert = 2,

        [Display(Name = "Directory Convert")]
        DirectoryConvert = 3,

        [Display(Name = "Quit")]
        Quit = 4
    }
}
