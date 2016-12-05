using System;
using System.Data;
using System.IO;
using System.Linq;
using System.Security.Cryptography;
using System.Text.RegularExpressions;
using System.Web;



public class Helper
{
    public static string[,] GetArrayFromIRange (SpreadsheetGear.IRange ir)
    {
        string[,] sValues = new string[ir.RowCount,ir.ColumnCount];
        for (int iCol = 0; iCol < ir.ColumnCount; iCol++)
        {
            for (int iRow = 0; iRow < ir.RowCount; iRow++)
            {
                sValues[iRow, iCol] = ir.Cells[iRow, iCol].Value.ToString();
            }
        }
        return sValues;
    }
/// <summary>
/// Gets a DataTable from an IRange assumingt the first row contains heading names.  All cells are strings.
/// Could eventually develp a directo JSON string equivolent... but that can come later.
/// </summary>
/// <param name="ir">Single area IRange with headers in first row</param>
/// <returns></returns>
    public static DataTable GetTableFromIRange(SpreadsheetGear.IRange ir)
    {
        DataTable dt = new DataTable();
        for (int iCol = 0; iCol < ir.ColumnCount; iCol++)
        {
            dt.Columns.Add(ir.Cells[0, iCol].Value.ToString(), typeof(string));
        }
        for (int iRow = 1; iRow < ir.RowCount; iRow++)
        {
            DataRow row = dt.NewRow();
            for (int iCol = 0; iCol < ir.ColumnCount; iCol++)
            {
                row[iCol] = ir.Cells[iRow, iCol].Value.ToString();
            }
            dt.Rows.Add(row);
        }
        return dt;
                 
    }

    public static string GetMD5Hash (byte[] b)
    {
        MD5 md5 = MD5.Create();
        byte[] byteHash = md5.ComputeHash(b);
        return string.Concat(byteHash.Select(x => x.ToString("X2")));
    }
    
    /// <summary>
    /// Reads data from a stream until the end is reached. The
    /// data is returned as a byte array. An IOException is
    /// thrown if any of the underlying IO calls fail.
    /// </summary>
    /// <param name="stream">The stream to read data from</param>
    /// <param name="initialLength">The initial buffer length</param>
    public static byte[] ReadFully(System.IO.Stream stream, long initialLength)
    {
        // reset pointer just in case
        stream.Seek(0, System.IO.SeekOrigin.Begin);

        // If we've been passed an unhelpful initial length, just
        // use 32K.
        if (initialLength < 1)
        {
            initialLength = 32768;
        }

        byte[] buffer = new byte[initialLength];
        int read = 0;

        int chunk;
        while ((chunk = stream.Read(buffer, read, buffer.Length - read)) > 0)
        {
            read += chunk;

            // If we've reached the end of our buffer, check to see if there's
            // any more information
            if (read == buffer.Length)
            {
                int nextByte = stream.ReadByte();

                // End of stream? If so, we're done
                if (nextByte == -1)
                {
                    return buffer;
                }

                // Nope. Resize the buffer, put in the byte we've just
                // read, and continue
                byte[] newBuffer = new byte[buffer.Length * 2];
                Array.Copy(buffer, newBuffer, buffer.Length);
                newBuffer[read] = (byte)nextByte;
                buffer = newBuffer;
                read++;
            }
        }
        // Buffer is now too big. Shrink it.
        byte[] ret = new byte[read];
        Array.Copy(buffer, ret, read);
        return ret;
    }
}

public static class StringExtensions
{
    /// <summary>
    /// takes a substring between two anchor strings (or the end of the string if that anchor is null)
    /// </summary>
    /// <param name="this">a string</param>
    /// <param name="from">an optional string to search after</param>
    /// <param name="until">an optional string to search before</param>
    /// <param name="comparison">an optional comparison for the search</param>
    /// <returns>a substring based on the search</returns>
    public static string Substring(this string @this, string from = null, string until = null, StringComparison comparison = StringComparison.InvariantCulture)
    {
        var fromLength = (from ?? string.Empty).Length;
        var startIndex = !string.IsNullOrEmpty(from)
            ? @this.IndexOf(from, comparison) + fromLength
            : 0;

        if (startIndex < fromLength) { throw new ArgumentException("from: Failed to find an instance of the first anchor"); }

        var endIndex = !string.IsNullOrEmpty(until)
        ? @this.IndexOf(until, startIndex, comparison)
        : @this.Length;

        if (endIndex < 0) { throw new ArgumentException("until: Failed to find an instance of the last anchor"); }

        var subString = @this.Substring(startIndex, endIndex - startIndex);
        return subString;
    }
}

