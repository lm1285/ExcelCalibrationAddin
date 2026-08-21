using System;
using System.Diagnostics;
using System.Threading.Tasks;
using ExcelCalibrationAddin.Host.UseCases;

namespace ExcelCalibrationAddin.Vsto
{
    internal sealed class DailySyncScheduler
    {
        private readonly Func<Task<TemplateSyncRunResult>> _sync;
        private readonly Func<DateTime?> _lastSyncUtc;
        private bool _running;

        public DailySyncScheduler(Func<Task<TemplateSyncRunResult>> sync, Func<DateTime?> lastSyncUtc)
        {
            _sync = sync ?? throw new ArgumentNullException(nameof(sync));
            _lastSyncUtc = lastSyncUtc ?? throw new ArgumentNullException(nameof(lastSyncUtc));
        }

        public async Task<TemplateSyncRunResult> RunIfDueAsync()
        {
            if (_running || !IsDue())
            {
                return null;
            }

            _running = true;
            try
            {
                var result = await _sync();
                if (result != null && !result.Succeeded)
                {
                    Trace.WriteLine($"[VSTO] Daily template sync failed: {result.ErrorMessage}");
                }

                return result;
            }
            catch (Exception ex)
            {
                Trace.WriteLine($"[VSTO] Daily template sync failed: {ex}");
                return new TemplateSyncRunResult { ErrorMessage = ex.Message };
            }
            finally
            {
                _running = false;
            }
        }

        private bool IsDue()
        {
            var lastSync = _lastSyncUtc();
            return !lastSync.HasValue || DateTime.UtcNow - lastSync.Value >= TimeSpan.FromDays(1);
        }
    }
}
