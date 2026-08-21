using System;
using System.Collections.Generic;
using System.Data.SQLite;
using System.Linq;
using ExcelCalibrationAddin.Core.Models;
using Newtonsoft.Json;

namespace ExcelCalibrationAddin.Core.Repositories
{
    public sealed partial class LocalTemplateRuleCacheRepository
    {
        public long SaveSampleDataVersion(string templateFingerprint, IReadOnlyList<TemplateSampleData> items, string remark = null)
        {
            if (string.IsNullOrWhiteSpace(templateFingerprint)) throw new ArgumentException("模板指纹不能为空。", nameof(templateFingerprint));
            var safeItems = (items ?? new List<TemplateSampleData>()).Where(item => item != null && !string.IsNullOrWhiteSpace(item.CalibrationItemName)).ToList();
            if (safeItems.Count == 0) throw new InvalidOperationException("没有可保存的样本数据。");
            using (var connection = OpenConnection())
            using (var transaction = connection.BeginTransaction())
            {
                var now = DateTime.UtcNow;
                long versionId;
                using (var command = connection.CreateCommand())
                {
                    command.Transaction = transaction;
                    command.CommandText = "INSERT INTO SampleDataVersion(TemplateFingerprint,CreatedAt,Remark,SyncStatus) VALUES(@fp,@at,@remark,0); SELECT last_insert_rowid();";
                    command.Parameters.AddWithValue("@fp", templateFingerprint);
                    command.Parameters.AddWithValue("@at", now.ToString("o"));
                    command.Parameters.AddWithValue("@remark", remark ?? string.Empty);
                    versionId = Convert.ToInt64(command.ExecuteScalar());
                }
                foreach (var item in safeItems)
                {
                    long itemId;
                    using (var command = connection.CreateCommand())
                    {
                        command.Transaction = transaction;
                        command.CommandText = "INSERT INTO TemplateSampleData(VersionId,CalibrationItemName,CalibrationItemKey) VALUES(@version,@name,@key); SELECT last_insert_rowid();";
                        command.Parameters.AddWithValue("@version", versionId);
                        command.Parameters.AddWithValue("@name", item.CalibrationItemName.Trim());
                        command.Parameters.AddWithValue("@key", item.CalibrationItemKey ?? string.Empty);
                        itemId = Convert.ToInt64(command.ExecuteScalar());
                    }
                    foreach (var point in item.Points ?? new List<SampleDataPoint>())
                    {
                        if (point == null || point.MeasurementValues == null || point.MeasurementValues.Count == 0 || point.MeasurementValues.Count > 30) continue;
                        using (var command = connection.CreateCommand())
                        {
                            command.Transaction = transaction;
                            command.CommandText = @"INSERT INTO SampleDataPoint(SampleDataId,PointIndex,SourceRow,SourceColumn,StandardValue,MeasurementValues,DecimalPlaces)
VALUES(@item,@index,@row,@column,@standard,@values,@decimals);";
                            command.Parameters.AddWithValue("@item", itemId);
                            command.Parameters.AddWithValue("@index", point.PointIndex);
                            command.Parameters.AddWithValue("@row", point.SourceRow);
                            command.Parameters.AddWithValue("@column", point.SourceColumn);
                            command.Parameters.AddWithValue("@standard", (object)point.StandardValue ?? DBNull.Value);
                            command.Parameters.AddWithValue("@values", JsonConvert.SerializeObject(point.MeasurementValues));
                            command.Parameters.AddWithValue("@decimals", Math.Max(0, Math.Min(15, point.DecimalPlaces)));
                            command.ExecuteNonQuery();
                        }
                    }
                }
                transaction.Commit();
                return versionId;
            }
        }

        public IReadOnlyList<SampleDataVersion> ListSampleDataVersions(string templateFingerprint)
        {
            var result = new List<SampleDataVersion>();
            using (var connection = OpenConnection())
            using (var command = connection.CreateCommand())
            {
                command.CommandText = "SELECT Id,TemplateFingerprint,CreatedAt,Remark,SyncStatus FROM SampleDataVersion WHERE TemplateFingerprint=@fp ORDER BY CreatedAt DESC;";
                command.Parameters.AddWithValue("@fp", templateFingerprint ?? string.Empty);
                using (var reader = command.ExecuteReader())
                    while (reader.Read()) result.Add(new SampleDataVersion { Id = Convert.ToInt64(reader[0]), TemplateFingerprint = Convert.ToString(reader[1]), CreatedAt = ParseDate(reader[2]), Remark = Convert.ToString(reader[3]), SyncStatus = Convert.ToInt32(reader[4]) });
            }
            foreach (var version in result) LoadSampleItems(version);
            return result;
        }

        public bool DeleteSampleDataVersion(long versionId)
        {
            using (var connection = OpenConnection()) using (var command = connection.CreateCommand())
            {
                command.CommandText = "DELETE FROM SampleDataVersion WHERE Id=@id;";
                command.Parameters.AddWithValue("@id", versionId);
                return command.ExecuteNonQuery() > 0;
            }
        }

        public IReadOnlyList<SampleDataPoint> ListLatestSampleDataPoints(string templateFingerprint)
        {
            var versions = ListSampleDataVersions(templateFingerprint);
            var result = new List<SampleDataPoint>();
            var seen = new HashSet<string>(StringComparer.Ordinal);
            foreach (var version in versions.OrderByDescending(item => item.CreatedAt))
            {
                foreach (var item in version.Items ?? new List<TemplateSampleData>())
                {
                    foreach (var point in item.Points ?? new List<SampleDataPoint>())
                    {
                        var key = (item.CalibrationItemKey ?? item.CalibrationItemName ?? string.Empty) + "|" + point.PointIndex;
                        if (seen.Add(key)) { point.CalibrationItemName = item.CalibrationItemName; result.Add(point); }
                    }
                }
            }
            return result;
        }

        private void LoadSampleItems(SampleDataVersion version)
        {
            using (var connection = OpenConnection()) using (var command = connection.CreateCommand())
            {
                command.CommandText = "SELECT Id,CalibrationItemName,CalibrationItemKey FROM TemplateSampleData WHERE VersionId=@version ORDER BY Id;";
                command.Parameters.AddWithValue("@version", version.Id);
                using (var reader = command.ExecuteReader())
                    while (reader.Read()) version.Items.Add(new TemplateSampleData { Id = Convert.ToInt64(reader[0]), VersionId = version.Id, CalibrationItemName = Convert.ToString(reader[1]), CalibrationItemKey = Convert.ToString(reader[2]) });
            }
            foreach (var item in version.Items)
            {
                using (var connection = OpenConnection()) using (var command = connection.CreateCommand())
                {
                    command.CommandText = "SELECT Id,PointIndex,SourceRow,SourceColumn,StandardValue,MeasurementValues,DecimalPlaces FROM SampleDataPoint WHERE SampleDataId=@item ORDER BY PointIndex;";
                    command.Parameters.AddWithValue("@item", item.Id);
                    using (var reader = command.ExecuteReader()) while (reader.Read()) item.Points.Add(new SampleDataPoint { Id = Convert.ToInt64(reader[0]), SampleDataId = item.Id, PointIndex = Convert.ToInt32(reader[1]), SourceRow = Convert.ToInt32(reader[2]), SourceColumn = Convert.ToInt32(reader[3]), StandardValue = reader.IsDBNull(4) ? (double?)null : Convert.ToDouble(reader[4]), MeasurementValues = DeserializeValues(Convert.ToString(reader[5])), DecimalPlaces = Convert.ToInt32(reader[6]) });
                }
            }
        }

        private static DateTime ParseDate(object value) { DateTime result; return DateTime.TryParse(Convert.ToString(value), null, System.Globalization.DateTimeStyles.RoundtripKind, out result) ? result : DateTime.MinValue; }
        private static List<double> DeserializeValues(string raw) { try { return JsonConvert.DeserializeObject<List<double>>(raw) ?? new List<double>(); } catch { return new List<double>(); } }
    }
}
