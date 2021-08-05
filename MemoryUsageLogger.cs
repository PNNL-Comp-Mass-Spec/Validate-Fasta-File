// *********************************************************************************************************
// Written by Matthew Monroe for the US Department of Energy
// Pacific Northwest National Laboratory, Richland, WA
// Created 02/09/2009
// Last updated 02/03/2016
// *********************************************************************************************************

using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;

namespace ValidateFastaFile
{
    /// <summary>
    /// Memory usage logger
    /// </summary>
    public class MemoryUsageLogger
    {
        // Ignore Spelling: yyyy-MM-dd, hh:mm:ss tt, nonpaged

        // The minimum interval between appending a new memory usage entry to the log
        private float mMinimumMemoryUsageLogIntervalMinutes = 1;

        // Used to determine the amount of free memory
        private PerformanceCounter mPerfCounterFreeMemory;
        private PerformanceCounter mPerfCounterPoolPagedBytes;
        private PerformanceCounter mPerfCounterPoolNonpagedBytes;

        private bool mPerfCountersInitialized;

        private readonly List<string> mHeaderNames;

        private readonly List<int> mHeaderNameLengths;

        /// <summary>
        /// Output folder for the log file
        /// </summary>
        /// <remarks>If this is an empty string, the log file is created in the working directory</remarks>
        public string LogFolderPath { get; }

        /// <summary>
        /// The minimum interval between appending a new memory usage entry to the log
        /// </summary>
        public float MinimumLogIntervalMinutes
        {
            get => mMinimumMemoryUsageLogIntervalMinutes;
            set
            {
                if (value < 0)
                    value = 0;
                mMinimumMemoryUsageLogIntervalMinutes = value;
            }
        }

        /// <summary>
        /// Constructor
        /// </summary>
        /// <param name="logFolderPath">Folder in which to write the memory log file(s); if this is an empty string, the log file is created in the working directory</param>
        /// <param name="minLogIntervalMinutes">Minimum log interval, in minutes</param>
        /// <remarks>
        /// Use WriteMemoryUsageLogEntry to append an entry to the log file.
        /// Alternatively use GetMemoryUsageSummary() to retrieve the memory usage as a string</remarks>
        public MemoryUsageLogger(string logFolderPath, float minLogIntervalMinutes = 5)
        {
            if (string.IsNullOrWhiteSpace(logFolderPath))
            {
                LogFolderPath = string.Empty;
            }
            else
            {
                LogFolderPath = logFolderPath;
            }

            MinimumLogIntervalMinutes = minLogIntervalMinutes;

            mHeaderNames = new List<string>
            {
                "Date".PadRight(10),
                "Time".PadRight(11),
                "ProcessMemoryUsage_MB",
                "FreeMemory_MB",
                "PoolPaged_MB",
                "PoolNonpaged_MB"
            };

            mHeaderNameLengths = new List<int>();
            foreach (var item in mHeaderNames)
            {
                mHeaderNameLengths.Add(item.Length + 2);
            }
        }

        /// <summary>
        /// Returns the amount of free memory on the current machine
        /// </summary>
        /// <returns>Free memory, in MB</returns>
        public float GetFreeMemoryMB()
        {
            try
            {
                if (mPerfCounterFreeMemory == null)
                {
                    return 0;
                }

                return mPerfCounterFreeMemory.NextValue();
            }
            catch (Exception)
            {
                // Ignore errors here
                return -1;
            }
        }

        /// <summary>
        /// Return the memory usage columns as a space or tab-separated list
        /// </summary>
        /// <param name="tabSeparated"></param>
        public string GetMemoryUsageHeader(bool tabSeparated = false)
        {
            return GetFormattedValues(mHeaderNames, mHeaderNameLengths, tabSeparated);
        }

        /// <summary>
        /// Get memory usage data as a space or tab-separated list
        /// </summary>
        /// <param name="tabSeparated"></param>
        public string GetMemoryUsageSummary(bool tabSeparated = false)
        {
            if (!mPerfCountersInitialized)
            {
                InitializePerfCounters();
            }

            var currentTime = DateTime.Now;

            var dataValues = new List<string>
            {
                currentTime.ToString("yyyy-MM-dd"),
                currentTime.ToString("hh:mm:ss tt"),
                GetProcessMemoryUsageMB().ToString("0.0"),
                GetFreeMemoryMB().ToString("0.0"),
                GetPoolPagedMemory().ToString("0.0"),
                GetPoolNonpagedMemory().ToString("0.0")
            };

            return GetFormattedValues(dataValues, mHeaderNameLengths, tabSeparated);
        }

        /// <summary>
        /// Returns the amount of pool nonpaged memory on the current machine
        /// </summary>
        /// <returns>Pool Nonpaged memory, in MB</returns>
        public float GetPoolNonpagedMemory()
        {
            try
            {
                if (mPerfCounterPoolNonpagedBytes == null)
                {
                    return 0;
                }
                else
                {
                    return (float)(mPerfCounterPoolNonpagedBytes.NextValue() / 1024.0 / 1024);
                }
            }
            catch (Exception)
            {
                // Ignore errors here
                return -1;
            }
        }

        /// <summary>
        /// Returns the amount of pool paged memory on the current machine
        /// </summary>
        /// <returns>Pool Paged memory, in MB</returns>
        public float GetPoolPagedMemory()
        {
            try
            {
                if (mPerfCounterPoolPagedBytes == null)
                {
                    return 0;
                }

                return (float)(mPerfCounterPoolPagedBytes.NextValue() / 1024.0 / 1024);
            }
            catch (Exception)
            {
                // Ignore errors here
                return -1;
            }
        }

        /// <summary>
        /// Returns the amount of memory that the currently running process is using
        /// </summary>
        /// <returns>Memory usage, in MB</returns>
        public static float GetProcessMemoryUsageMB()
        {
            try
            {
                // Obtain a handle to the current process
                var objProcess = Process.GetCurrentProcess();

                // The WorkingSet is the total physical memory usage
                return (float)(objProcess.WorkingSet64 / 1024.0 / 1024);
            }
            catch (Exception)
            {
                // Ignore errors here
                return 0;
            }
        }

        /// <summary>
        /// Initializes the performance counters
        /// </summary>
        /// <returns>Any errors that occur; empty string if no errors</returns>
        public string InitializePerfCounters()
        {
            var msgErrors = string.Empty;

            try
            {
                mPerfCounterFreeMemory = new PerformanceCounter("Memory", "Available MBytes") { ReadOnly = true };
            }
            catch (Exception ex)
            {
                if (msgErrors.Length > 0)
                    msgErrors += "; ";
                msgErrors += "Error instantiating the Memory: 'Available MBytes' performance counter: " + ex.Message;
            }

            try
            {
                mPerfCounterPoolPagedBytes = new PerformanceCounter("Memory", "Pool Paged Bytes") { ReadOnly = true };
            }
            catch (Exception ex)
            {
                if (msgErrors.Length > 0)
                    msgErrors += "; ";
                msgErrors += "Error instantiating the Memory: 'Pool Paged Bytes' performance counter: " + ex.Message;
            }

            try
            {
                mPerfCounterPoolNonpagedBytes =
                    new PerformanceCounter("Memory", "Pool NonPaged Bytes") { ReadOnly = true };
            }
            catch (Exception ex)
            {
                if (msgErrors.Length > 0)
                    msgErrors += "; ";
                msgErrors += "Error instantiating the Memory: 'Pool NonPaged Bytes' performance counter: " + ex.Message;
            }

            mPerfCountersInitialized = true;

            return msgErrors;
        }

        private DateTime dtLastWriteTime = DateTime.UtcNow.Subtract(TimeSpan.FromHours(1));

        private string GetFormattedValues(
            IReadOnlyList<string> values,
            IReadOnlyList<int> headerNameLengths,
            bool tabSeparated = false)
        {
            if (tabSeparated)
            {
                var trimmedValues = values.Select(item => item.Trim()).ToList();
                return string.Join("\t", trimmedValues);
            }

            var dataLine = new StringBuilder();

            for (var i = 0; i < headerNameLengths.Count; i++)
            {
                dataLine.Append(values[i].PadRight(headerNameLengths[i]));
            }

            return dataLine.ToString();
        }

        /// <summary>
        /// Writes a status file tracking memory usage
        /// </summary>
        // ReSharper disable once UnusedMember.Global
        public void WriteMemoryUsageLogEntry()
        {
            try
            {
                if (DateTime.UtcNow.Subtract(dtLastWriteTime).TotalMinutes < mMinimumMemoryUsageLogIntervalMinutes)
                {
                    // Not enough time has elapsed since the last write; exit sub
                    return;
                }

                dtLastWriteTime = DateTime.UtcNow;

                // We're creating a new log file each month
                var logFileName = "MemoryUsageLog_" + DateTime.Now.ToString("yyyy-MM") + ".txt";

                string logFilePath;
                if (!string.IsNullOrWhiteSpace(LogFolderPath))
                {
                    logFilePath = Path.Combine(LogFolderPath, logFileName);
                }
                else
                {
                    logFilePath = string.Copy(logFileName);
                }

                var writeHeader = !File.Exists(logFilePath);

                using var writer = new StreamWriter(new FileStream(logFilePath, FileMode.Append, FileAccess.Write, FileShare.Read));

                if (writeHeader)
                {
                    GetMemoryUsageHeader(true);
                }

                writer.WriteLine(GetMemoryUsageSummary(true));
            }
            catch
            {
                // Ignore errors here
            }
        }
    }
}