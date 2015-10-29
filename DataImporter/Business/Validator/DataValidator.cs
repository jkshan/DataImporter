using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using DataImporter.Models;

namespace DataImporter.Business.Validator
{
    public static class DataValidator
    {
        /// <summary>
        /// Loads the Valid List of ISO Currency Codes.
        /// </summary>
        static Lazy<List<string>> CurrencyCodes = new Lazy<List<string>>(() =>
        {
            return (from c in CultureInfo.GetCultures(CultureTypes.SpecificCultures)
                    let r = new RegionInfo(c.LCID)
                    select r.ISOCurrencySymbol).ToList();
        });

        /// <summary>
        /// Validates the Records for Errors
        /// </summary>
        /// <param name="row">Row which needs to be validated</param>
        /// <returns>Returns the Row State(Valid or Invalid)</returns>
        public static bool ValidateData(Record row)
        {
            if(string.IsNullOrWhiteSpace(row.Account))
            {
                row.ErrorMessages = "Account Column is Empty";
            }

            if(string.IsNullOrWhiteSpace(row.Description))
            {
                row.ErrorMessages = string.Join(",", row.ErrorMessages, "Description Column is Empty");
            }

            if(string.IsNullOrWhiteSpace(row.CurrencyCode))
            {
                row.ErrorMessages = string.Join(",", row.ErrorMessages, "Currency Code Column is Empty");
            }

            if(string.IsNullOrWhiteSpace(row.Amount))
            {
                row.ErrorMessages = string.Join(",", row.ErrorMessages, "Amount Column is Empty");
            }

            decimal temp;
            if(!decimal.TryParse(row.Amount, out temp))
            {
                row.ErrorMessages = string.Join(",", row.ErrorMessages, "Amount must be a Decimal Value");
            }

            if(row.CurrencyCode.Length != 3 && CurrencyCodes.Value.FirstOrDefault(i => string.Equals(row.CurrencyCode, i, StringComparison.OrdinalIgnoreCase)) == null)
            {
                row.ErrorMessages = string.Join(",", row.ErrorMessages, "Invalide Currency Code");
            }

            if(!string.IsNullOrWhiteSpace(row.ErrorMessages))
            {
                row.ErrorMessages = row.ErrorMessages.TrimStart(',');
            }

            return string.IsNullOrWhiteSpace(row.ErrorMessages);
        }
    }
}