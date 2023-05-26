// See https://aka.ms/new-console-template for more information
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

class FinancialAnalysisFromExcel
{
    static void Main()
    {
        // Load the Excel file
        FileInfo file = new FileInfo("stock_prices_latest.xlsx");
        using (ExcelPackage package = new ExcelPackage(file))
        {
            ExcelWorksheet worksheet = package.Workbook.Worksheets["stock_prices_latest"]; // Assuming the data is on the first sheet (index 0)

            // Read the data from the Excel sheet
            List<List<double>> assetData = new List<List<double>>();
            int rowCount = worksheet.Dimension.Rows;
            int colCount = worksheet.Dimension.Columns;

            for (int row = 2; row <= rowCount; row++) // Assuming the data starts from the second row
            {
                List<double> rowData = new List<double>();
                for (int col = 3; col <= colCount; col++) // Assuming the data starts from the second column
                {
                    double value = double.Parse(worksheet.Cells[row, col].Value.ToString());
                    rowData.Add(value);
                }
                assetData.Add(rowData);
            }

            // Calculate returns
            List<List<double>> returns = CalculateReturns(assetData);

            // Calculate volatility
            List<double> volatilities = CalculateVolatility(returns);

            // Calculate correlation
            List<List<double>> correlationMatrix = CalculateCorrelationMatrix(returns);

            // Display the results
            Console.WriteLine("Returns:");
            DisplayMatrix(returns);

            Console.WriteLine("\nVolatilities:");
            DisplayList(volatilities);

            Console.WriteLine("\nCorrelation Matrix:");
            DisplayMatrix(correlationMatrix);
        }

        // Wait for user input to exit
        Console.ReadLine();
    }

    // Calculate returns for each asset
    static List<List<double>> CalculateReturns(List<List<double>> assetData)
    {
        List<List<double>> returns = new List<List<double>>();

        foreach (var asset in assetData)
        {
            List<double> assetReturns = new List<double>();

            for (int i = 1; i < asset.Count; i++)
            {
                double returnVal = (asset[i] - asset[i - 1]) / asset[i - 1];
                assetReturns.Add(returnVal);
            }

            returns.Add(assetReturns);
        }

        return returns;
    }

    // Calculate volatility for each asset
    static List<double> CalculateVolatility(List<List<double>> returns)
    {
        List<double> volatilities = new List<double>();

        foreach (var asset in returns)
        {
            double volatility = Math.Sqrt(asset.Average(r => Math.Pow(r - asset.Average(), 2)));
            volatilities.Add(volatility);
        }

        return volatilities;
    }

    // Calculate correlation matrix between assets
    static List<List<double>> CalculateCorrelationMatrix(List<List<double>> returns)
    {
        int assetCount = returns.Count;
        List<List<double>> correlationMatrix = new List<List<double>>();

        for (int i = 0; i < assetCount; i++)
        {
            List<double> row = new List<double>();

            for (int j = 0; j < assetCount; j++)
            {
                double correlation = CalculateCorrelation(returns[i], returns[j]);
                row.Add(correlation);
            }

            correlationMatrix.Add(row);
        }

        return correlationMatrix;
    }

    // Calculate correlation between two sets of returns
    static double CalculateCorrelation(List<double> returns1, List<double> returns2)
    {
        double mean1 = returns1.Average();
        double mean2 = returns2.Average();

        double sum = 0;
        int count = returns1.Count;

        for (int i = 0; i < count; i++)
        {
            sum += (returns1[i] - mean1) * (returns2[i] - mean2);
        }

        double covariance = sum / (count - 1);
        double stdDeviation1 = Math.Sqrt(returns1.Sum(r => Math.Pow(r - mean1, 2)) / (count - 1));
        double stdDeviation2 = Math.Sqrt(returns2.Sum(r => Math.Pow(r - mean2, 2)) / (count - 1));

        double correlation = covariance / (stdDeviation1 * stdDeviation2);
        return correlation;
    }

    // Display a matrix
    static void DisplayMatrix(List<List<double>> matrix)
    {
        foreach (var row in matrix)
        {
            foreach (var value in row)
            {
                Console.Write(value.ToString("0.000") + "\t");
            }
            Console.WriteLine();
        }
    }

    // Display a list
    static void DisplayList(List<double> list)
    {
        foreach (var value in list)
        {
            Console.Write(value.ToString("0.000") + "\t");
        }
        Console.WriteLine();
    }
}


