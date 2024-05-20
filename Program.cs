using System;
using System.Collections;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Numerics;
using System.Text;
using System.Timers;
using OfficeOpenXml;

public static class Project
{
    static string[] sample_cases = {
        @"D:\Algo\Algo project\Test Cases\Sample\1-Input.xlsx",
        @"D:\Algo\Algo project\Test Cases\Sample\2-Input.xlsx",
        @"D:\Algo\Algo project\Test Cases\Sample\3-Input.xlsx",
        @"D:\Algo\Algo project\Test Cases\Sample\4-Input.xlsx",
        @"D:\Algo\Algo project\Test Cases\Sample\5-Input.xlsx",
        @"D:\Algo\Algo project\Test Cases\Sample\6-Input.xlsx",
    };

    static string[] easy_cases = {
        @"D:\Algo\Algo project\Test Cases\Complete\Easy\1-Input.xlsx",
        @"D:\Algo\Algo project\Test Cases\Complete\Easy\2-Input.xlsx",
    };

    static string[] medium_cases = {
        @"D:\Algo\Algo project\Test Cases\Complete\Medium\1-Input.xlsx",
        @"D:\Algo\Algo project\Test Cases\Complete\Medium\2-Input.xlsx",
    };

    static string[] hard_cases = {
        @"D:\Algo\Algo project\Test Cases\Complete\Hard\1-Input.xlsx",
        @"D:\Algo\Algo project\Test Cases\Complete\Hard\2-Input.xlsx",
    };

    static Dictionary<long, long> parent;
    static Dictionary<long, int> size;
    //tuple<bigger percentage,lines matched,smaller percentage,file1 node,file2 node,file1 node hyperlink,file2 node hyperlink>
    static (int, int, int, long, long, string, string, Uri, Uri)[] edges;
    static List<(string, string, long, int, Uri, Uri)> mst_edges;
    static Dictionary<long, HashSet<long>> components;
    static Dictionary<long, float[]> avg;
    static Dictionary<long, List<KeyValuePair<long, long>>> adj;
    static Stopwatch mst_sw = new Stopwatch(), stat_sw = new Stopwatch(), total_sw = new Stopwatch();

    private static void init()
    {
        mst_edges = new List<(string, string, long, int, Uri, Uri)>();
        components = new Dictionary<long, HashSet<long>>();
        avg = new Dictionary<long, float[]>();
        parent = new Dictionary<long, long>();
        size = new Dictionary<long, int>();
        adj = new Dictionary<long, List<KeyValuePair<long, long>>>();

        foreach (var edge in edges)
        {
            if (!parent.ContainsKey(edge.Item4))
            {
                parent.Add(edge.Item4, edge.Item4);
                size.Add(edge.Item4, 1);
                adj.Add(edge.Item4, new List<KeyValuePair<long, long>>());
            }
            if (!parent.ContainsKey(edge.Item5))
            {
                parent.Add(edge.Item5, edge.Item5);
                size.Add(edge.Item5, 1);
                adj.Add(edge.Item5, new List<KeyValuePair<long, long>>());
            }
        }

        constructGraph();
    }

    private static long dsu_find(long node)
    {
        if (parent[node] == node) return node;
        return parent[node] = dsu_find(parent[node]);
    }

    private static bool dsu_union(long node1, long node2)
    {
        long par1 = dsu_find(node1), par2 = dsu_find(node2);
        if (par1 == par2) return false;

        if (size[par1] < size[par2])
        {
            parent[par1] = par2;
            size[par2] += size[par1];
        }
        else
        {
            parent[par2] = par1;
            size[par1] += size[par2];
        }
        return true;
    }

    private static void extractInfo(string edge1, string edge2, ref int percentage1, ref int percentage2, ref long id1, ref long id2)
    {
        bool f1 = true, f2 = true;

        foreach (var c in edge1)
        {
            if (c == '(') f1 = false;

            if (c >= '0' && c <= '9')
            {
                if (f1)
                {
                    id1 *= 10;
                    id1 += c - '0';
                }
                else
                {
                    percentage1 *= 10;
                    percentage1 += c - '0';
                }
            }
        }

        foreach (var c in edge2)
        {
            if (c == '(') f2 = false;

            if (c >= '0' && c <= '9')
            {
                if (f2)
                {
                    id2 *= 10;
                    id2 += c - '0';
                }
                else
                {
                    percentage2 *= 10;
                    percentage2 += c - '0';
                }
            }
        }
    }

    public static void read(string file_path)
    {
        using (ExcelPackage excelPackage1 = new ExcelPackage(new FileInfo(file_path)))
        {
            ExcelWorksheet worksheet1 = excelPackage1.Workbook.Worksheets[0];
            int rowCount = worksheet1.Dimension.Rows;
            edges = new (int, int, int, long, long, string, string, Uri, Uri)[rowCount - 1];

            for (int row = 2; row <= rowCount; row++)
            {
                string edge1 = worksheet1.Cells[row, 1].Value.ToString();
                string edge2 = worksheet1.Cells[row, 2].Value.ToString();
                Uri hyper_link1 = worksheet1.Cells[row, 1].Hyperlink;
                Uri hyper_link2 = worksheet1.Cells[row, 2].Hyperlink;
                int matched_lines = int.Parse(worksheet1.Cells[row, 3].Value.ToString());
                long id1 = 0, id2 = 0;
                int percentage1 = 0, percentage2 = 0;
                extractInfo(edge1, edge2, ref percentage1, ref percentage2, ref id1, ref id2);

                if (percentage1 > percentage2)
                    edges[row - 2] = (percentage1, matched_lines, percentage2, id1, id2, edge1, edge2, hyper_link1, hyper_link2);
                else edges[row - 2] = (percentage2, matched_lines, percentage1, id1, id2, edge1, edge2, hyper_link1, hyper_link2);
            }
        }
    }

    public static void constructGraph()
    {
        foreach (var i in edges)
        {
            adj[i.Item5].Add(new KeyValuePair<long, long>(i.Item4, i.Item1));
            adj[i.Item4].Add(new KeyValuePair<long, long>(i.Item5, i.Item1));
        }
    }

    public static void MST()
    {
        Array.Sort(edges);
        Array.Reverse(edges);
        foreach (var edge in edges)
        {
            if (dsu_union(edge.Item4, edge.Item5))
            {
                mst_edges.Add((edge.Item6, edge.Item7, edge.Item4, edge.Item2, edge.Item8, edge.Item9));
            }
        }
    }

    public static void calculateStat()
    {
        foreach (var edge in edges)
        {
            long par = dsu_find(edge.Item4);
            if (!components.ContainsKey(par))
            {
                components.Add(par, new HashSet<long>());
            }
            components[par].Add(edge.Item4);
            components[par].Add(edge.Item5);

            if (!avg.ContainsKey(par))
            {
                avg.Add(par, new float[2]);
                avg[par][0] = avg[par][1] = 0;
            }
            avg[par][0] += edge.Item1 + edge.Item3;
            avg[par][1] += 2;
        }
        //storing each component avg in idx 0 in the array
        foreach (var item in avg)
        {
            item.Value[0] /= item.Value[1];
        }
    }

    public static void writeStatFile(string directory, int cnt)
    {
        using (ExcelPackage excelPackage = new ExcelPackage())
        {
            (double, List<long>)[] stat = new (double, List<long>)[components.Count];
            int i = 0;
            foreach (var comp in components)
            {
                List<long> tmp = [.. comp.Value];
                tmp.Sort();
                stat[i++] = (Math.Round(avg[comp.Key][0], 1), tmp);
            }

            Array.Sort(stat, (x, y) =>
            {
                return x.Item1.CompareTo(y.Item1);
            });
            Array.Reverse(stat);

            ExcelWorksheet worksheet = excelPackage.Workbook.Worksheets.Add("Sheet1");
            int row = 1;

            worksheet.Cells[row, 1].Value = "Component Index";
            worksheet.Cells[row, 2].Value = "Vertices";
            worksheet.Cells[row, 3].Value = "Average Similarity";
            worksheet.Cells[row, 4].Value = "Component Count";

            foreach (var comp in stat)
            {
                ++row;
                string ids = "";
                int size = comp.Item2.Count;
                worksheet.Cells[row, 4].Value = size;

                foreach (var node in comp.Item2)
                {
                    ids += node.ToString();
                    if (--size > 0) ids += ", ";
                }

                worksheet.Cells[row, 1].Value = row - 1;
                worksheet.Cells[row, 2].Value = ids;
                worksheet.Cells[row, 3].Value = comp.Item1;
            }

            worksheet.Column(1).AutoFit();
            worksheet.Column(2).AutoFit();
            worksheet.Column(3).AutoFit();
            worksheet.Column(4).AutoFit();

            excelPackage.SaveAs(new FileInfo($@"D:\Algo\Algo project\{directory}\{cnt}-stat file{cnt}.xlsx"));
        }
    }

    public static void writeMstFile(string directory, int cnt)
    {
        using (ExcelPackage excelPackage = new ExcelPackage())
        {
            (float, long, int, string, string, Uri, Uri)[] mst_file = new (float, long, int, string, string, Uri, Uri)[mst_edges.Count];
            int i = 0;
            foreach (var edge in mst_edges)
            {
                mst_file[i++] = (avg[parent[edge.Item3]][0], parent[edge.Item3], edge.Item4, edge.Item1, edge.Item2, edge.Item5, edge.Item6);
            }

            Array.Sort(mst_file, (x, y) =>
            {
                int comp1 = x.Item1.CompareTo(y.Item1);
                if (comp1 == 0)
                {
                    int comp2 = x.Item2.CompareTo(y.Item2);
                    if (comp2 == 0)
                    {
                        return x.Item3.CompareTo(y.Item3);
                    }
                    return comp2;
                }
                return comp1;
            });
            Array.Reverse(mst_file);

            ExcelWorksheet worksheet = excelPackage.Workbook.Worksheets.Add("Sheet1");
            var hyperlinkStyle = worksheet.Workbook.Styles.CreateNamedStyle("hyperlink");
            hyperlinkStyle.Style.Font.UnderLine = true;
            hyperlinkStyle.Style.Font.Color.SetColor(Color.Blue);

            int row = 1;
            worksheet.Cells[row, 1].Value = "File 1";
            worksheet.Cells[row, 2].Value = "File 2";
            worksheet.Cells[row, 3].Value = "Line Matches";

            foreach (var tuple in mst_file)
            {
                ++row;
                worksheet.Cells[row, 1].Hyperlink = tuple.Item6;
                worksheet.Cells[row, 1].Value = tuple.Item4;
                worksheet.Cells[row, 1].StyleName = "hyperlink";

                worksheet.Cells[row, 2].Hyperlink = tuple.Item7;
                worksheet.Cells[row, 2].Value = tuple.Item5;
                worksheet.Cells[row, 2].StyleName = "hyperlink";

                worksheet.Cells[row, 3].Value = tuple.Item3;
            }

            worksheet.Column(1).AutoFit();
            worksheet.Column(2).AutoFit();
            worksheet.Column(3).AutoFit();

            excelPackage.SaveAs(new FileInfo($@"D:\Algo\Algo project\{directory}\{cnt}-mst file{cnt}.xlsx"));
        }
    }

    public static void sampleCases()
    {
        Console.WriteLine("Running sample cases\n");
        for (int i = 0; i < sample_cases.GetLength(0); ++i)
        {
            total_sw.Restart();

            read(sample_cases[i]);
            init();

            mst_sw.Restart();
            MST();
            mst_sw.Stop();

            stat_sw.Restart();
            calculateStat();
            stat_sw.Stop();

            mst_sw.Start();
            writeMstFile("sample cases output", i + 1);
            mst_sw.Stop();

            stat_sw.Start();
            writeStatFile("sample cases output", i + 1);
            stat_sw.Stop();

            total_sw.Stop();

            Console.WriteLine($"case {i + 1}:");
            Console.WriteLine($"stat time is {stat_sw.ElapsedMilliseconds} ms");
            Console.WriteLine($"mst time is {mst_sw.ElapsedMilliseconds} ms");
            Console.WriteLine($"total time is {total_sw.ElapsedMilliseconds} ms\n");
        }

        Console.WriteLine("\n");
    }

    public static void easyCases()
    {
        Console.WriteLine("Running complete easy cases\n");
        for (int i = 0; i < easy_cases.GetLength(0); ++i)
        {
            total_sw.Restart();

            read(easy_cases[i]);
            init();

            mst_sw.Restart();
            MST();
            mst_sw.Stop();

            stat_sw.Restart();
            calculateStat();
            stat_sw.Stop();

            mst_sw.Start();
            writeMstFile("easy cases output", i + 1);
            mst_sw.Stop();

            stat_sw.Start();
            writeStatFile("easy cases output", i + 1);
            stat_sw.Stop();

            total_sw.Stop();

            Console.WriteLine($"case {i + 1}:");
            Console.WriteLine($"stat time is {stat_sw.ElapsedMilliseconds} ms");
            Console.WriteLine($"mst time is {mst_sw.ElapsedMilliseconds} ms");
            Console.WriteLine($"total time is {total_sw.ElapsedMilliseconds} ms\n");
        }

        Console.WriteLine("\n");
    }

    public static void mediumCases()
    {
        Console.WriteLine("Running complete medium cases\n");
        for (int i = 0; i < medium_cases.GetLength(0); ++i)
        {
            total_sw.Restart();

            read(medium_cases[i]);
            init();

            mst_sw.Restart();
            MST();
            mst_sw.Stop();

            stat_sw.Restart();
            calculateStat();
            stat_sw.Stop();

            mst_sw.Start();
            writeMstFile("medium cases output", i + 1);
            mst_sw.Stop();

            stat_sw.Start();
            writeStatFile("medium cases output", i + 1);
            stat_sw.Stop();

            total_sw.Stop();

            Console.WriteLine($"case {i + 1}:");
            Console.WriteLine($"stat time is {stat_sw.ElapsedMilliseconds} ms");
            Console.WriteLine($"mst time is {mst_sw.ElapsedMilliseconds} ms");
            Console.WriteLine($"total time is {total_sw.ElapsedMilliseconds} ms\n");
        }

        Console.WriteLine("\n");
    }

    public static void hardCases()
    {
        Console.WriteLine("Running complete hard cases\n");
        for (int i = 0; i < hard_cases.GetLength(0); ++i)
        {
            total_sw.Restart();

            read(hard_cases[i]);
            init();

            mst_sw.Restart();
            MST();
            mst_sw.Stop();

            stat_sw.Restart();
            calculateStat();
            stat_sw.Stop();

            mst_sw.Start();
            writeMstFile("hard cases output", i + 1);
            mst_sw.Stop();

            stat_sw.Start();
            writeStatFile("hard cases output", i + 1);
            stat_sw.Stop();

            total_sw.Stop();

            Console.WriteLine($"case {i + 1}:");
            Console.WriteLine($"stat time is {stat_sw.ElapsedMilliseconds} ms");
            Console.WriteLine($"mst time is {mst_sw.ElapsedMilliseconds} ms");
            Console.WriteLine($"total time is {total_sw.ElapsedMilliseconds} ms\n");
        }
    }

    public static void Main()
    {
        ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
        sampleCases();
        easyCases();
        mediumCases();
        hardCases();
    }
}
