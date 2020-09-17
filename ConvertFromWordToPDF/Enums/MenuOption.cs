using System.ComponentModel.DataAnnotations;

namespace ConvertFromWordToPDF.Enums
{
    public enum MenuOption
    {
        [Display(Name = "Single Convert")]
        SingleConvert = 1,

        [Display(Name = "Directory Convert")]
        DirectoryConvert = 2,

        [Display(Name = "Quit")]
        Quit = 3
    }
}
