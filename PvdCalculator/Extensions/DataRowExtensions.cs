using System;
using System.Data;

namespace PvdCalculator.Extensions
{
    public static class DataRowExtensions
    {
        public static T Field<T>(this DataRow row, int columnIndex)
        {
            if (row != null)
            {
                if (!row.IsNull(columnIndex) && !string.IsNullOrEmpty(row[columnIndex].ToString()))
                {
                    Type t = typeof(T);
                    if (t == typeof(bool) || t == typeof(bool?))
                    {
                        switch (row[columnIndex].ToString().ToUpper())
                        {
                            case "1":
                            case "TRUE":
                            case "Y":
                            case "YES":
                                return (T)(object)true;
                            case "0":
                            case "FALSE":
                            case "N":
                            case "NO":
                                return (T)(object)false;
                            default:
                                return default(T);
                        }
                    }

                    return (T)Convert.ChangeType(row[columnIndex], Nullable.GetUnderlyingType(t) ?? t);
                }

                return default(T);
            }

            throw new ArgumentNullException(nameof(row));
        }

        public static T Field<T>(this DataRow row, string columnName)
        {
            if (row != null)
            {
                if (row.Table.Columns.Contains(columnName) && !row.IsNull(columnName))
                {
                    Type t = typeof(T);
                    if (t == typeof(bool) || t == typeof(bool?))
                    {
                        switch (row[columnName].ToString().ToUpper())
                        {
                            case "1":
                            case "TRUE":
                            case "Y":
                            case "YES":
                                return (T)(object)true;
                            case "0":
                            case "FALSE":
                            case "N":
                            case "NO":
                                return (T)(object)false;
                            default:
                                return default(T);
                        }
                    }

                    return (T)Convert.ChangeType(row[columnName], Nullable.GetUnderlyingType(t) ?? t);
                }

                return default(T);
            }

            throw new ArgumentNullException(nameof(row));
        }
    }
}
