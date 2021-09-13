using NPOI.SS.UserModel;
using System;
using System.Collections.Generic;

namespace NpoiMapper
{
    public static class TypeExtensions
    {
        /// <summary>
        /// Collection of numeric types.
        /// </summary>
        private static readonly HashSet<Type> NumericTypes = new HashSet<Type>
        {
            typeof(decimal),
            typeof(byte), typeof(sbyte),
            typeof(short), typeof(ushort),
            typeof(int), typeof(uint),
            typeof(long), typeof(ulong),
            typeof(float), typeof(double)
        };

        /// <summary>
        /// Check if the given type is a numeric type.
        /// </summary>
        /// <param name="type">The type to be checked.</param>
        /// <returns><c>true</c> if it's numeric; otherwise <c>false</c>.</returns>
        public static bool IsNumeric(this Type type)
        {
            return NumericTypes.Contains(type);
        }

        public static bool AsBool(this object value)
        {
            return Convert.ToBoolean(value);
        }
        public static string AsString(this object value)
        {
            return value as string;
        }
        public static int AsInt(this object value)
        {
            return Convert.ToInt32(value);
        }
        public static uint AsUInt(this object value)
        {
            return Convert.ToUInt32(value);
        }
        public static float AsFloat(this object value)
        {
            return float.Parse(value.AsString());
        }
        public static double AsDouble(this object value)
        {
            return Convert.ToDouble(value);
        }
        public static byte AsByte(this object value)
        {
            return Convert.ToByte(value);
        }
        public static sbyte AsSByte(this object value)
        {
            return Convert.ToSByte(value);
        }
        public static short AsShort(this object value)
        {
            return Convert.ToInt16(value);
        }
        public static ushort AsUShort(this object value)
        {
            return Convert.ToUInt16(value);
        }
        public static long AsLong(this object value)
        {
            return Convert.ToInt64(value);
        }
        public static ulong AsULong(this object value)
        {
            return Convert.ToUInt64(value);
        }
        public static decimal AsDecimal(this object value)
        {
            return Convert.ToDecimal(value);
        }
        public static DateTime AsDateTime(this object value)
        {
            if (!(value is double))
                return Convert.ToDateTime(value);
            else
                return DateUtil.GetJavaDate((double)value, false);
        }
        public static IEnumerable<T> AsMapSource<T>(this IEnumerable<T> value)
        {
            var source = value as IEnumerable<T>;
            if (source != null)
                return source;
            return new T[0];
        }
    }
}