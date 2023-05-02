
namespace ExcelColumnTransformer
{
    internal class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("Please input an excel column (ex. V, BC, ABC): ");
            var columnInput = Console.ReadLine();
            var columnNumberOutput = ExcelColumnNameToNumber(columnInput);
            Console.WriteLine($"The number of the {columnInput} excel column is {columnNumberOutput}");


            Console.WriteLine("Please input an number: ");
            var numberInput = Console.ReadLine();
            var excelColumnOutput = NumberToExcelColumnName(Convert.ToInt32(numberInput));
            Console.WriteLine($"The excel column of the number {numberInput} is {excelColumnOutput}");

            Console.ReadLine();
        }

        public static int ExcelColumnNameToNumber(string input)
        {
            var result = 0;
            foreach (var character in input.ToUpper())
            {
                result *= 26;
                result += character - 'A' + 1;
            }

            return result;
        }

        public static string NumberToExcelColumnName(int columnNumber)
        {
            string columnName = "";
            while (columnNumber > 0)
            {
                int remainder = (columnNumber - 1) % 26;
                columnName = (char)('A' + remainder) + columnName;
                columnNumber = (columnNumber - remainder) / 26;
            }

            return columnName;
        }
    }
}