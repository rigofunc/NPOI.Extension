using System.ComponentModel.DataAnnotations;
using System.Reflection;
using FluentExcel;

namespace samples
{
    public static class FluentConfigurationExtensions
    {
        public static FluentConfiguration<TModel> FromAnnotations<TModel>(this FluentConfiguration<TModel> fluentConfiguration) where TModel : class
        {
            var properties = typeof(TModel).GetProperties();
            foreach (var property in properties)
            {
                var pc = fluentConfiguration.Property(property);

                var display = property.GetCustomAttribute<DisplayAttribute>();
                if (display != null)
                {
                    pc.HasExcelTitle(display.Name);
                    if (display.GetOrder().HasValue)
                    {
                        pc.HasExcelIndex(display.Order);
                    }
                }
                else
                {
                    pc.HasExcelTitle(property.Name);
                }

                var format = property.GetCustomAttribute<DisplayFormatAttribute>();
                if (format != null)
                {
                    pc.HasDataFormatter(format.DataFormatString
                                              .Replace("{0:", "")
                                              .Replace("}", ""));
                }

                if (pc.Index < 0)
                {
                    pc.HasAutoIndex();
                }
            }

            return fluentConfiguration;
        }
    }
}
