using OfficeOpenXml;

namespace MS_8laba;

public class Program
{
    public static void Main()
    {
        var BASE_PARAMS = new Params(
            k: 1,
            l: 8,
            m: 2,
            n: 7,
            k1: 90,
            b: 25000,
            i1: 10,
            i2: 2,
            s: 200,
            V: 800,
            T: 12,
            deltaMax: 0.5
        );

        var factors = new List<Factor>(4)
        {
            new Factor(0.8, 1.2),
            new Factor(6.4, 9.6),
            new Factor(1.6, 2.4),
            new Factor(5.6, 8.4)
        };

        var experements = new List<ExperementItem>();
        var placements = GeneratePlacements(2, 4);

        for (int j = 0; j < placements.Count; j++)
        {
            var placement = placements[j];
            var variant = new List<double>();
            for (int i = 0; i < placement.Count; i++)
            {
                var factor = factors[i];
                if (placement[i] == 0)
                {
                    variant.Add(factor.MinValue);
                }
                else if (placement[i] == 1)
                {
                    variant.Add(factor.MaxValue);
                }
            }

            var param = BASE_PARAMS with { k = variant[0], l = variant[1], m = variant[2], n = variant[3] };
            var result = Start(0.01, param);

            var x = result.x;
            var y = result.ys.Select(y => y[2]).ToList();

            double s = 0.0;
            for (int i = 0; i < x.Count - 1; i++)
            {
                double x1 = Math.Abs(x[i]);
                double x2 = Math.Abs(x[i + 1]);
                double y1 = Math.Abs(y[i]);
                double y2 = Math.Abs(y[i + 1]);
                s += (y1 + y2) / 2 * (x2 - x1);
            }

            experements.Add(new ExperementItem(j + 1, param, s));
        }

        // shuffle
        var rnd = new Random();
        for (int i = 0; i < experements.Count - 1; i++)
        {
            int j = rnd.Next(i + 1, experements.Count);
            (experements[i], experements[j]) = (experements[j], experements[i]);
        }

        var cwd = new DirectoryInfo(Environment.CurrentDirectory);
        var pathToBaseFolder = cwd?.Parent?.Parent?.Parent?.FullName;
        var pathToExcelFile = Path.Combine(pathToBaseFolder!, "ms_8laba.xlsx");
        var worksheetName = "Regressions";

        ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
        var package = new ExcelPackage(pathToExcelFile);
        var sheet = package.Workbook.Worksheets[worksheetName];

        // START: ('J' 28) END: ('O' 43)
        var range = sheet.Cells[28, 10, 43, 15];
        for (int i = 0; i < experements.Count; i++)
        {
            range.SetCellValue(i, 0, experements[i].Num);
            range.SetCellValue(i, 1, experements[i].Params.k);
            range.SetCellValue(i, 2, experements[i].Params.l);
            range.SetCellValue(i, 3, experements[i].Params.m);
            range.SetCellValue(i, 4, experements[i].Params.n);
            range.SetCellValue(i, 5, experements[i].S);
        }

        File.WriteAllBytes(pathToExcelFile, package.GetAsByteArray());
    }

    public static Result Start(double h, Params param)
    {
        List<double> times = new() { 0.0 };
        List<List<double>> values = new() { new List<double>() { 0.3, 0.3, 0, 0, 500 } };
        List<double> deltaValues = new() { values[0][3] };

        var f0 = (double t, int j) => param.k * values[j][1] - param.k * values[j][0];

        var f1 = (double t, int j) => values[j][2];

        var f2 = (double t, int j) => param.l * values[j][0]
                        - param.l * values[j][1]
                        - param.m * values[j][2]
                        + param.n * values[j][3];

        var f3 = (double nextDelta) => Math.Abs(nextDelta) <= param.deltaMax
            ? nextDelta
            : param.deltaMax * Math.Sign(nextDelta);

        var f4 = (double t, int j) => param.V * Math.Sin(values[j][0]);

        var omega = (double t, int j) => (10000 - values[j][4]) / (param.b - param.V * t);

        var delta = (double t, int j) => -param.k1 * values[j][3]
                            - param.i1 * values[j][1]
                            - param.i2 * values[j][2]
                            + param.s * (omega(t, j) - values[j][1]);

        var eulerMethod = (double prevValue, int prevStep, Func<double, int, double> f, double t) => {
            return prevValue + f(t, prevStep) * h;
        };

        var solve = (int nextStep, double t) => {
            var newValues = new List<double>();
            var prevStep = nextStep - 1;

            newValues.Add(eulerMethod(values[prevStep][0], prevStep, f0, t));
            newValues.Add(eulerMethod(values[prevStep][1], prevStep, f1, t));
            newValues.Add(eulerMethod(values[prevStep][2], prevStep, f2, t));

            var nextDelta = eulerMethod(deltaValues[prevStep], prevStep, delta, t);
            deltaValues.Add(nextDelta);

            newValues.Add(f3(nextDelta));
            newValues.Add(eulerMethod(values[prevStep][4], prevStep, f4, t));

            return newValues;
        };

        double t = 0.0;
        int j = 0;
        while (t < param.T)
        {
            t += h;
            j += 1;
            var nextValues = solve(j, t);
            times.Add(t);
            values.Add(nextValues);
        }

        return new Result(x: times, ys: values);
    }

    public static List<List<int>> GeneratePlacements(int n, int m)
    {
        var h = Math.Max(n, m);
        var x = Enumerable.Range(1, h).Select(_ => 0).ToList();
        var all = new List<List<int>> { x.Select(x => x).ToList() };

        while (true)
        {
            var j = m - 1;
            while (j >= 0 && x[j] == n - 1)
            {
                j--;
            }

            if (j < 0)
            {
                break;
            }

            if (x[j] >= n)
            {
                j--;
            }

            x[j]++;

            if (j == m - 1)
            {
                all.Add(x.Select(x => x).ToList());
                continue;
            }

            for (int k = j + 1; k < m; k++)
            {
                x[k] = 0;
            }

            all.Add(x.Select(x => x).ToList());
        }

        return all;
    }
}

public record ExperementItem(
    int Num,
    Params Params, 
    double S);

public record Result(
    List<double> x, 
    List<List<double>> ys);

public record Params(
    double k, 
    double l, 
    double m, 
    double n, 
    double k1, 
    double b, 
    double i1, 
    double i2, 
    double s, 
    double V, 
    double T, 
    double deltaMax);

public record Factor(
    double MinValue,
    double MaxValue);