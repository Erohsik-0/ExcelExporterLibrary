using ExportExcel.Interfaces;
using ExportExcel.Models;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Globalization;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using System.Text.RegularExpressions;
using ValidationResult = ExportExcel.Models.ValidationResult;

namespace ExportExcel.Services
{
    /// <summary>
    /// Data validation service for Excel operations
    /// </summary>
    public class DataValidator : IDataValidator
    {
        private readonly ValidationLevel _validationLevel;
        private readonly Dictionary<string, Func<object, bool>> _customValidators;
        private readonly HashSet<string> _requiredFields;
        private readonly Dictionary<string, Type> _fieldTypes;

        public DataValidator(ValidationLevel validationLevel = ValidationLevel.Standard)
        {
            _validationLevel = validationLevel;
            _customValidators = new Dictionary<string, Func<object, bool>>(StringComparer.OrdinalIgnoreCase);
            _requiredFields = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            _fieldTypes = new Dictionary<string, Type>(StringComparer.OrdinalIgnoreCase);
            InitializeDefaultValidators();
        }

        public ValidationResult ValidateData(List<Dictionary<string, object>> data)
        {
            var result = new ValidationResult();
            if (data == null)
            {
                result.AddError("Data cannot be null");
                return result;
            }

            // Basic validation
            if (_validationLevel >= ValidationLevel.Basic)
            {
                ValidateBasicDataStructure(data, result);
            }

            // Standard validation
            if (_validationLevel >= ValidationLevel.Standard)
            {
                ValidateDataConsistency(data, result);
            }

            // Strict validation
            if (_validationLevel >= ValidationLevel.Strict)
            {
                ValidateDataIntegrity(data, result);
            }

            // Validate individual records
            for (int i = 0; i < data.Count; i++)
            {
                var recordResult = ValidateRecord(data[i], i);
                if (!recordResult.IsValid)
                {
                    foreach (var error in recordResult.Errors)
                    {
                        result.AddError($"Record {i + 1}: {error}");
                    }
                    foreach (var warning in recordResult.Warnings)
                    {
                        result.AddWarning($"Record {i + 1}: {warning}");
                    }
                }
            }

            // Add metadata
            result.AddMetadata("TotalRecords", data.Count);
            result.AddMetadata("ValidationLevel", _validationLevel.ToString());
            result.AddMetadata("ValidatedAt", DateTime.UtcNow);
            return result;
        }

        public ValidationResult ValidateRecord(Dictionary<string, object> record, int recordIndex)
        {
            var result = new ValidationResult();
            if (record == null)
            {
                result.AddError("Record cannot be null");
                return result;
            }

            // Basic validation
            if (_validationLevel >= ValidationLevel.Basic)
            {
                ValidateRequiredFields(record, result);
                ValidateFieldTypes(record, result);
            }

            // Standard validation
            if (_validationLevel >= ValidationLevel.Standard)
            {
                ValidateDataFormats(record, result);
                ValidateBusinessRules(record, result);
            }

            // Strict validation
            if (_validationLevel >= ValidationLevel.Strict)
            {
                ValidateCustomRules(record, result);
                ValidateDataRelationships(record, result);
            }
            return result;
        }

        public ValidationResult ValidateHeaders(IEnumerable<string> headers)
        {
            var result = new ValidationResult();
            var headerList = headers?.ToList() ?? new List<string>();
            if (headerList.Count == 0)
            {
                result.AddError("Headers cannot be empty");
                return result;
            }

            // Check for duplicate headers
            var duplicates = headerList.GroupBy(h => h, StringComparer.OrdinalIgnoreCase)
                                       .Where(g => g.Count() > 1)
                                       .Select(g => g.Key);
            foreach (var duplicate in duplicates)
            {
                result.AddError($"Duplicate header found: '{duplicate}'");
            }

            // Check for invalid header names
            var invalidHeaderPattern = new Regex(@"^[a-zA-Z_][a-zA-Z0-9_]*$");
            foreach (var header in headerList)
            {
                if (string.IsNullOrWhiteSpace(header))
                {
                    result.AddError("Empty or whitespace header found");
                    continue;
                }
                if (_validationLevel >= ValidationLevel.Standard)
                {
                    if (!invalidHeaderPattern.IsMatch(header.Replace(".", "_").Replace("[", "_").Replace("]", "_")))
                    {
                        result.AddWarning($"Header '{header}' contains special characters that might cause issues");
                    }
                }
            }

            // Check for required headers
            if (_validationLevel >= ValidationLevel.Standard)
            {
                var missingRequired = _requiredFields.Except(headerList, StringComparer.OrdinalIgnoreCase);
                foreach (var missing in missingRequired)
                {
                    result.AddError($"Required header missing: '{missing}'");
                }
            }

            result.AddMetadata("HeaderCount", headerList.Count);
            result.AddMetadata("UniqueHeaders", headerList.Distinct(StringComparer.OrdinalIgnoreCase).Count());
            return result;
        }

        #region Configuration Methods
        public void AddRequiredField(string fieldName)
        {
            if (!string.IsNullOrWhiteSpace(fieldName))
            {
                _requiredFields.Add(fieldName);
            }
        }

        public void AddFieldType(string fieldName, Type expectedType)
        {
            if (!string.IsNullOrWhiteSpace(fieldName) && expectedType != null)
            {
                _fieldTypes[fieldName] = expectedType;
            }
        }

        public void AddCustomValidator(string fieldName, Func<object, bool> validator)
        {
            if (!string.IsNullOrWhiteSpace(fieldName) && validator != null)
            {
                _customValidators[fieldName] = validator;
            }
        }
        #endregion

        #region Private Validation Methods
        private void InitializeDefaultValidators()
        {
            // Common email validation
            _customValidators["email"] = value =>
            {
                if (value == null) return true;
                return new EmailAddressAttribute().IsValid(value.ToString());
            };

            // Common phone validation
            _customValidators["phone"] = value =>
            {
                if (value == null) return true;
                return new PhoneAttribute().IsValid(value.ToString());
            };

            // Common URL validation
            _customValidators["url"] = value =>
            {
                if (value == null) return true;
                return new UrlAttribute().IsValid(value.ToString());
            };
        }

        private void ValidateBasicDataStructure(List<Dictionary<string, object>> data, ValidationResult result)
        {
            if (data.Count == 0)
            {
                result.AddWarning("Data list is empty");
                return;
            }

            // Check for consistent structure
            var firstRecord = data[0];
            var expectedKeys = new HashSet<string>(firstRecord.Keys, StringComparer.OrdinalIgnoreCase);
            for (int i = 1; i < data.Count; i++)
            {
                var currentKeys = new HashSet<string>(data[i].Keys, StringComparer.OrdinalIgnoreCase);
                if (!currentKeys.SetEquals(expectedKeys))
                {
                    result.AddWarning($"Record {i + 1} has a different column structure than the first record.");
                }
            }
        }

        private void ValidateDataConsistency(List<Dictionary<string, object>> data, ValidationResult result)
        {
            if (data.Count == 0) return;

            var fieldStats = new Dictionary<string, FieldStatistics>(StringComparer.OrdinalIgnoreCase);

            // Collect statistics for each field
            foreach (var record in data)
            {
                foreach (var kvp in record)
                {
                    if (!fieldStats.ContainsKey(kvp.Key))
                    {
                        fieldStats[kvp.Key] = new FieldStatistics();
                    }
                    var stats = fieldStats[kvp.Key];
                    stats.TotalCount++;
                    if (kvp.Value == null || string.IsNullOrWhiteSpace(kvp.Value.ToString()))
                    {
                        stats.NullOrEmptyCount++;
                    }
                    else
                    {
                        var valueType = kvp.Value.GetType();
                        if (!stats.ObservedTypes.ContainsKey(valueType))
                        {
                            stats.ObservedTypes[valueType] = 0;
                        }
                        stats.ObservedTypes[valueType]++;
                    }
                }
            }

            // Check for inconsistencies
            foreach (var kvp in fieldStats)
            {
                var fieldName = kvp.Key;
                var stats = kvp.Value;

                // High null percentage warning
                var nullPercentage = (double)stats.NullOrEmptyCount / stats.TotalCount;
                if (nullPercentage > 0.8)
                {
                    result.AddWarning($"Field '{fieldName}' has {nullPercentage:P0} null or empty values.");
                }

                // Mixed type warning
                if (stats.ObservedTypes.Count > 1)
                {
                    var types = string.Join(", ", stats.ObservedTypes.Keys.Select(t => t.Name));
                    result.AddWarning($"Field '{fieldName}' has mixed data types: {types}.");
                }
            }
        }

        private void ValidateDataIntegrity(List<Dictionary<string, object>> data, ValidationResult result)
        {
            // Check for duplicate records
            var recordHashes = new HashSet<string>();
            var duplicateIndices = new List<int>();
            for (int i = 0; i < data.Count; i++)
            {
                var hash = GenerateRecordHash(data[i]);
                if (!recordHashes.Add(hash))
                {
                    duplicateIndices.Add(i + 1);
                }
            }
            if (duplicateIndices.Any())
            {
                result.AddWarning($"Found {duplicateIndices.Count} potential duplicate records at rows: {string.Join(", ", duplicateIndices)}");
            }

            // Check for referential integrity if ID fields are present
            ValidateReferentialIntegrity(data, result);
        }

        private void ValidateRequiredFields(Dictionary<string, object> record, ValidationResult result)
        {
            foreach (var requiredField in _requiredFields)
            {
                if (!record.TryGetValue(requiredField, out var value) || value == null || string.IsNullOrWhiteSpace(value.ToString()))
                {
                    result.AddError($"Required field '{requiredField}' is missing or empty.");
                }
            }
        }

        private void ValidateFieldTypes(Dictionary<string, object> record, ValidationResult result)
        {
            foreach (var kvp in _fieldTypes)
            {
                var fieldName = kvp.Key;
                var expectedType = kvp.Value;
                if (record.TryGetValue(fieldName, out var value) && value != null)
                {
                    try
                    {
                        // Attempt to convert to the target type to validate
                        Convert.ChangeType(value, expectedType);
                    }
                    catch (Exception)
                    {
                        result.AddError($"Field '{fieldName}' with value '{value}' cannot be converted to the expected type {expectedType.Name}.");
                    }
                }
            }
        }

        private void ValidateDataFormats(Dictionary<string, object> record, ValidationResult result)
        {
            foreach (var kvp in record)
            {
                var fieldName = kvp.Key.ToLower();
                var value = kvp.Value;
                if (value == null) continue;
                var stringValue = value.ToString();

                // Email validation
                if (fieldName.Contains("email") && !string.IsNullOrEmpty(stringValue))
                {
                    if (!new EmailAddressAttribute().IsValid(stringValue))
                    {
                        result.AddError($"Invalid email format in field '{kvp.Key}': {stringValue}");
                    }
                }

                // Date validation
                if (fieldName.Contains("date") || fieldName.Contains("time"))
                {
                    if (value is not DateTime && !DateTime.TryParse(stringValue, out _))
                    {
                        result.AddWarning($"Potential invalid date format in field '{kvp.Key}': {stringValue}");
                    }
                }

                // Numeric validation for price/amount fields
                if (fieldName.Contains("price") || fieldName.Contains("amount") || fieldName.Contains("cost"))
                {
                    if (!decimal.TryParse(stringValue, NumberStyles.Any, CultureInfo.InvariantCulture, out var numericValue))
                    {
                        result.AddWarning($"Invalid numeric format in field '{kvp.Key}': {stringValue}");
                    }
                    else if (numericValue < 0)
                    {
                        result.AddWarning($"Negative value found in financial field '{kvp.Key}': {numericValue}");
                    }
                }
            }
        }

        private void ValidateBusinessRules(Dictionary<string, object> record, ValidationResult result)
        {
            // Example business rules - these would be customized based on domain
            // Age validation
            if (record.TryGetValue("age", out var ageValue) && ageValue != null)
            {
                if (int.TryParse(ageValue.ToString(), out var age))
                {
                    if (age < 0 || age > 150)
                    {
                        result.AddError($"Invalid age: {age}. Age must be between 0 and 150.");
                    }
                }
            }

            // Status consistency
            if (record.TryGetValue("status", out var statusValue) && record.TryGetValue("isActive", out var isActiveValue))
            {
                var status = statusValue?.ToString()?.ToLower();
                if (status == "inactive" && Convert.ToBoolean(isActiveValue))
                {
                    result.AddWarning("Inconsistent status: record is marked as 'inactive' but 'isActive' is true.");
                }
            }

            // Date range validation
            if (record.TryGetValue("startDate", out var startDateValue) && record.TryGetValue("endDate", out var endDateValue))
            {
                if (DateTime.TryParse(startDateValue?.ToString(), out var start) &&
                    DateTime.TryParse(endDateValue?.ToString(), out var end))
                {
                    if (start > end)
                    {
                        result.AddError("Start date cannot be after end date.");
                    }
                }
            }
        }

        private void ValidateCustomRules(Dictionary<string, object> record, ValidationResult result)
        {
            foreach (var kvp in record)
            {
                var fieldName = kvp.Key;
                var value = kvp.Value;

                // Check for exact field name match
                if (_customValidators.ContainsKey(fieldName))
                {
                    try
                    {
                        if (!_customValidators[fieldName](value))
                        {
                            result.AddError($"Custom validation failed for field '{fieldName}'.");
                        }
                    }
                    catch (Exception ex)
                    {
                        result.AddError($"Custom validation for field '{fieldName}' threw an exception: {ex.Message}");
                    }
                }

                // Check for field name patterns (e.g., a validator for "email" should also check "user_email")
                foreach (var validator in _customValidators)
                {
                    if (fieldName.ToLower().Contains(validator.Key.ToLower()) && validator.Key != fieldName)
                    {
                        try
                        {
                            if (!validator.Value(value))
                            {
                                result.AddWarning($"Pattern-based validation failed for field '{fieldName}' (using pattern: '{validator.Key}').");
                            }
                        }
                        catch (Exception ex)
                        {
                            result.AddWarning($"Pattern-based validation for field '{fieldName}' (pattern: '{validator.Key}') threw an exception: {ex.Message}");
                        }
                    }
                }
            }
        }

        private void ValidateDataRelationships(Dictionary<string, object> record, ValidationResult result)
        {
            // Example: Conditional Requirement
            // If 'country' is 'USA', then 'state' is required.
            if (record.TryGetValue("country", out var countryValue) && "usa".Equals(countryValue?.ToString(), StringComparison.OrdinalIgnoreCase))
            {
                if (!record.TryGetValue("state", out var stateValue) || stateValue == null || string.IsNullOrWhiteSpace(stateValue.ToString()))
                {
                    result.AddError("Field 'state' is required when 'country' is 'USA'.");
                }
            }
        }

        private void ValidateReferentialIntegrity(List<Dictionary<string, object>> data, ValidationResult result)
        {
            // Example placeholder for checking parent-child relationships
            // This is a complex topic and would require a more sophisticated setup
            if (data.Any(d => d.ContainsKey("id")) && data.Any(d => d.ContainsKey("parentId")))
            {
                var allIds = new HashSet<object>(data.Select(d => d["id"]).Where(id => id != null));
                for (int i = 0; i < data.Count; i++)
                {
                    var record = data[i];
                    if (record.TryGetValue("parentId", out var parentId) && parentId != null)
                    {
                        if (!allIds.Contains(parentId))
                        {
                            result.AddWarning($"Record {i + 1} has a 'parentId' ({parentId}) that does not exist in the dataset.");
                        }
                    }
                }
            }
        }

        private string GenerateRecordHash(Dictionary<string, object> record)
        {
            // Order by key to ensure consistency
            var orderedRecord = new SortedDictionary<string, object>(record, StringComparer.OrdinalIgnoreCase);
            var jsonString = JsonConvert.SerializeObject(orderedRecord);
            using (var sha256 = SHA256.Create())
            {
                var bytes = sha256.ComputeHash(Encoding.UTF8.GetBytes(jsonString));
                return Convert.ToBase64String(bytes);
            }
        }

        /// <summary>
        /// Helper class to hold statistics about fields during consistency validation.
        /// </summary>
        private class FieldStatistics
        {
            public int TotalCount { get; set; }
            public int NullOrEmptyCount { get; set; }
            public Dictionary<Type, int> ObservedTypes { get; } = new Dictionary<Type, int>();
        }

        #endregion
    }

}
