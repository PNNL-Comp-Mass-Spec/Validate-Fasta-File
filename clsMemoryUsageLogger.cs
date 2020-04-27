// *********************************************************************************************************
// Written by Matthew Monroe for the US Department of Energy
// Pacific Northwest National Laboratory, Richland, WA
// Created 02/09/2009
// Last updated 02/03/2016
// *********************************************************************************************************

using System;
using System.Diagnostics;
using System.IO;

namespace ValidateFastaFile
{
    public class clsMemoryUsageLogger
    {
        #region "Module variables"

        private const char COL_SEP = '\t';

        // The minimum interval between appending a new memory usage entry to the log
        private float m_MinimumMemoryUsageLogIntervalMinutes = 1;

        // Used to determine the amount of free memory
        private PerformanceCounter m_PerfCounterFreeMemory;
        private PerformanceCounter m_PerfCounterPoolPagedBytes;
        private PerformanceCounter m_PerfCounterPoolNonpagedBytes;

        private bool m_PerfCountersInitialized = false;
        #endregion

        #region "Properties"

        /// <summary>
        /// Output folder for the log file
        /// </summary>
        /// <value></value>
        /// <returns></returns>
        /// <remarks>If this is an empty string, the log file is created in the working directory</remarks>
        public string LogFolderPath { get; }

        /// <summary>
        /// The minimum interval between appending a new memory usage entry to the log
        /// </summary>
        /// <value></value>
        /// <returns></returns>
        /// <remarks></remarks>
        public float MinimumLogIntervalMinutes
        {
            get => m_MinimumMemoryUsageLogIntervalMinutes;
            set
            {
                if (value < 0)
                    value = 0;
                m_MinimumMemoryUsageLogIntervalMinutes = value;
            }
        }
        #endregion

        #region "Methods"

        /// <summary>
        /// Constructor
        /// </summary>
        /// <param name="logFolderPath">Folder in which to write the memory log file(s); if this is an empty string, the log file is created in the working directory</param>
        /// <param name="minLogIntervalMinutes">Minimum log interval, in minutes</param>
        /// <remarks>
        /// Use WriteMemoryUsageLogEntry to append an entry to the log file.
        /// Alternatively use GetMemoryUsageSummary() to retrieve the memory usage as a string</remarks>
        public clsMemoryUsageLogger(string logFolderPath, float minLogIntervalMinutes = 5)
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
        }

        /// <summary>
        /// Returns the amount of free memory on the current machine
        /// </summary>
        /// <returns>Free memory, in MB</returns>
        /// <remarks></remarks>
        public float GetFreeMemoryMB()
        {
            try
            {
                if (m_PerfCounterFreeMemory == null)
                {
                    return 0;
                }
                else
                {
                    return m_PerfCounterFreeMemory.NextValue();
                }
            }
            catch (Exception ex)
            {
                return -1;
            }
        }

        public string GetMemoryUsageHeader()
        {
            return "Date" + COL_SEP +
                   "Time" + COL_SEP +
                   "ProcessMemoryUsage_MB" + COL_SEP +
                   "FreeMemory_MB" + COL_SEP +
                   "PoolPaged_MB" + COL_SEP +
                   "PoolNonpaged_MB";
        }

        public string GetMemoryUsageSummary()
        {
            if (!m_PerfCountersInitialized)
            {
                InitializePerfCounters();
            }

            var currentTime = DateTime.Now;

            return currentTime.ToString("yyyy-MM-dd") + COL_SEP +
                   currentTime.ToString("hh:mm:ss tt") + COL_SEP +
                   GetProcessMemoryUsageMB().ToString("0.0") + COL_SEP +
                   GetFreeMemoryMB().ToString("0.0") + COL_SEP +
                   GetPoolPagedMemory().ToString("0.0") + COL_SEP +
                   GetPoolNonpagedMemory().ToString("0.0");
        }

        /// <summary>
        /// Returns the amount of pool nonpaged memory on the current machine
        /// </summary>
        /// <returns>Pool Nonpaged memory, in MB</returns>
        /// <remarks></remarks>
        public float GetPoolNonpagedMemory()
        {
            try
            {
                if (m_PerfCounterPoolNonpagedBytes == null)
                {
                    return 0;
                }
                else
                {
                    return (float)(m_PerfCounterPoolNonpagedBytes.NextValue() / 1024.0 / 1024);
                }
            }
            catch (Exception ex)
            {
                return -1;
            }
        }

        /// <summary>
        /// Returns the amount of pool paged memory on the current machine
        /// </summary>
        /// <returns>Pool Paged memory, in MB</returns>
        /// <remarks></remarks>
        public float GetPoolPagedMemory()
        {
            try
            {
                if (m_PerfCounterPoolPagedBytes == null)
                {
                    return 0;
                }
                else
                {
                    return (float)(m_PerfCounterPoolPagedBytes.NextValue() / 1024.0 / 1024);
                }
            }
            catch (Exception ex)
            {
                return -1;
            }
        }

        /// <summary>
        /// Returns the amount of memory that the currently running process is using
        /// </summary>
        /// <returns>Memory usage, in MB</returns>
        /// <remarks></remarks>
        public static float GetProcessMemoryUsageMB()
        {
            try
            {
                // Obtain a handle to the current process
                var objProcess = Process.GetCurrentProcess();

                // The WorkingSet is the total physical memory usage
                return (float)(objProcess.WorkingSet64 / 1024.0 / 1024);
            }
            catch (Exception ex)
            {
                return 0;
            }
        }

        /// <summary>
        /// Initializes the performance counters
        /// </summary>
        /// <returns>Any errors that occur; empty string if no errors</returns>
        /// <remarks></remarks>
        public string InitializePerfCounters()
        {
            string msgErrors = string.Empty;

            try
            {
                m_PerfCounterFreeMemory = new PerformanceCounter("Memory", "Available MBytes");
                m_PerfCounterFreeMemory.ReadOnly = true;
            }
            catch (Exception ex)
            {
                if (msgErrors.Length > 0)
                    msgErrors += "; ";
                msgErrors += "Error instantiating the Memory: 'Available MBytes' performance counter: " + ex.Message;
            }

            try
            {
                m_PerfCounterPoolPagedBytes = new PerformanceCounter("Memory", "Pool Paged Bytes");
                m_PerfCounterPoolPagedBytes.ReadOnly = true;
            }
            catch (Exception ex)
            {
                if (msgErrors.Length > 0)
                    msgErrors += "; ";
                msgErrors += "Error instantiating the Memory: 'Pool Paged Bytes' performance counter: " + ex.Message;
            }

            try
            {
                m_PerfCounterPoolNonpagedBytes = new PerformanceCounter("Memory", "Pool NonPaged Bytes");
                m_PerfCounterPoolNonpagedBytes.ReadOnly = true;
            }
            catch (Exception ex)
            {
                if (msgErrors.Length > 0)
                    msgErrors += "; ";
                msgErrors += "Error instantiating the Memory: 'Pool NonPaged Bytes' performance counter: " + ex.Message;
            }

            m_PerfCountersInitialized = true;

            return msgErrors;
        }

        private DateTime dtLastWriteTime = DateTime.UtcNow.Subtract(TimeSpan.FromHours(1));

        /// <summary>
        /// Writes a status file tracking memory usage
        /// </summary>
        /// <remarks></remarks>
        public void WriteMemoryUsageLogEntry()
        {
            try
            {
                if (DateTime.UtcNow.Subtract(dtLastWriteTime).TotalMinutes < m_MinimumMemoryUsageLogIntervalMinutes)
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

                bool writeHeader = !File.Exists(logFilePath);

                using (var writer = new StreamWriter(new FileStream(logFilePath, FileMode.Append, FileAccess.Write, FileShare.Read)))
                {
                    if (writeHeader)
                    {
                        GetMemoryUsageHeader();
                    }

                    writer.WriteLine(GetMemoryUsageSummary());
                }
            }
            catch
            {
                // Ignore errors here
            }
        }

        #endregion
    }
}