// Current project folder (one level above bin)
string projectRoot = Path.Combine(AppContext.BaseDirectory, "..", "..", "..");

// Full path to your template
var templatePath = Path.Combine(projectRoot, "Templates", "BaseShoppingList.ott");

// Output folder in project
var outputFolder = Path.Combine(projectRoot, "Output");

var service = new WordService(templatePath, outputFolder);

var replacements = new Dictionary<string, string>
{
    { "{{Date}}", DateTime.Now.ToShortDateString() },
    { "{{CustomerName}}", "Broccoli Mann" },
    { "{{Item1}}", "Apples" },
    { "{{Qty1}}", "5" },
    { "{{Item2}}", "Bananas" },
    { "{{Qty2}}", "3" }
};

service.Generate(replacements, "ShoppingList.odt");