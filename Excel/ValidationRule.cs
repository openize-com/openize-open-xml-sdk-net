using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml;

namespace Openize.Cells
{
    /// <summary>
    /// Specifies the types of validation that can be applied to a cell or range of cells.
    /// </summary>
    public enum ValidationType
    {
        /// <summary>Specifies a list validation type where cell value must be one of a predefined list.</summary>
        List,

        /// <summary>Specifies a date validation type where cell value must be a date within a specified range.</summary>
        Date,

        /// <summary>Specifies a whole number validation type where cell value must be a whole number within a specified range.</summary>
        WholeNumber,

        /// <summary>Specifies a decimal number validation type where cell value must be a decimal number within a specified range.</summary>
        Decimal,

        /// <summary>Specifies a text length validation type where the length of the cell text must be within a specified range.</summary>
        TextLength,

        /// <summary>Specifies a custom formula validation type where cell value must satisfy a custom formula.</summary>
        CustomFormula
    }

    /// <summary>
    /// Represents a validation rule that can be applied to a cell or range of cells.
    /// </summary>
    public class ValidationRule
    {
        /// <summary>Gets or sets the type of validation.</summary>
        public ValidationType Type { get; set; }

        /// <summary>Gets or sets the list of options for list validation.</summary>
        public string[] Options { get; set; }

        /// <summary>Gets or sets the minimum value for numeric or text length validation.</summary>
        public double? MinValue { get; set; }

        /// <summary>Gets or sets the maximum value for numeric or text length validation.</summary>
        public double? MaxValue { get; set; }

        /// <summary>Gets or sets the custom formula for custom formula validation.</summary>
        public string CustomFormula { get; set; }

        /// <summary>Gets or sets the title of the error dialog that appears when validation fails.</summary>
        public string ErrorTitle { get; set; } = "Invalid Input";

        /// <summary>Gets or sets the error message displayed when validation fails.</summary>
        public string ErrorMessage { get; set; } = "The value entered is invalid.";

        /// <summary>
        /// Initializes a new instance of the <see cref="ValidationRule"/> class for list validation.
        /// </summary>
        /// <param name="options">The options for the list validation.</param>
        public ValidationRule(string[] options)
        {
            Type = ValidationType.List;
            Options = options;
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="ValidationRule"/> class for numeric (whole number, decimal) or text length validations.
        /// </summary>
        /// <param name="type">The type of validation (whole number, decimal, text length).</param>
        /// <param name="minValue">The minimum value for the validation.</param>
        /// <param name="maxValue">The maximum value for the validation.</param>
        /// <exception cref="ArgumentException">Thrown if the type is not numeric or text length.</exception>
        // Constructor for numeric (decimal, whole number) and text length validation rules
        public ValidationRule(ValidationType type, double minValue, double maxValue)
        {
            if (type != ValidationType.WholeNumber && type != ValidationType.Decimal && type != ValidationType.TextLength)
            {
                throw new ArgumentException("This constructor is only for numeric and text length validations.");
            }

            Type = type;
            MinValue = minValue;
            MaxValue = maxValue;
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="ValidationRule"/> class for custom formula validation.
        /// </summary>
        /// <param name="customFormula">The custom formula for the validation.</param>
        public ValidationRule(string customFormula)
        {
            Type = ValidationType.CustomFormula;
            CustomFormula = customFormula;
        }

        // Optional: You can add more constructors or methods to initialize different types of validation rules
    }

}
