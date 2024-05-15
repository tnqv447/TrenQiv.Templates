using System.ComponentModel.DataAnnotations;
using System.Text.RegularExpressions;

namespace TrenQiv.Templates.Attributes;

public class RegularExpressionListAttribute(string pattern) : RegularExpressionAttribute(pattern)
{
    public override bool IsValid(object? value)
    {
        if (value is null)
        {
            return true; //null list
        }

        if (value is not IEnumerable<string> values)
        {
            return false;
        }

        foreach (var val in values)
        {
            if (!Regex.IsMatch(val, Pattern))
            {
                return false;
            }
        }

        return true;
    }
}
