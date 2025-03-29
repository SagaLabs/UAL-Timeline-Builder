"use client"

import type React from "react"

import { useState, useMemo, useRef, useEffect } from "react"
import Papa from "papaparse"
import {
  AlertTriangle,
  Calendar,
  Check,
  ChevronDown,
  Clock,
  Download,
  Filter,
  Globe,
  Hash,
  Info,
  Loader2,
  MapPin,
  RefreshCw,
  Search,
  Shield,
  User,
  X,
} from "lucide-react"

export default function UALTimelineBuilder() {
  const [logs, setLogs] = useState<any[]>([])
  const [fileName, setFileName] = useState("")
  const [userFilters, setUserFilters] = useState<string[]>([])
  const [workloadFilters, setWorkloadFilters] = useState<string[]>([])
  const [operationFilters, setOperationFilters] = useState<string[]>([])
  const [correlationFilter, setCorrelationFilter] = useState("")
  const [showOnlyRisky, setShowOnlyRisky] = useState(false)
  const [visibleCount, setVisibleCount] = useState(100)
  const [loading, setLoading] = useState(false)
  const [searchTerm, setSearchTerm] = useState("")

  const [userDropdownOpen, setUserDropdownOpen] = useState(false)
  const [workloadDropdownOpen, setWorkloadDropdownOpen] = useState(false)
  const [operationDropdownOpen, setOperationDropdownOpen] = useState(false)

  const userDropdownRef = useRef<HTMLDivElement>(null)
  const workloadDropdownRef = useRef<HTMLDivElement>(null)
  const operationDropdownRef = useRef<HTMLDivElement>(null)

  const riskyOps = [
    "UpdateInboxRule",
    "New-InboxRule",
    "Add-MailboxPermission",
    "Set-Mailbox",
    "Set-MailboxAutoReplyConfiguration",
    "Add member to role.",
    "Add user.",
    "Add delegated permission grant.",
    "Set-AdminAuditLogConfig",
    "Update application - Certificates and secrets management",
    "Consent to application.",
    "Add service principal.",
    "Update application.",
    "Add application.",
    "Add application permission.",
    "Update PasswordProfile.",
    "Change user password.",
    "Add owner to application."
  ]

  // Handle click outside to close dropdowns
  useEffect(() => {
    function handleClickOutside(event: MouseEvent) {
      if (userDropdownRef.current && !userDropdownRef.current.contains(event.target as Node)) {
        setUserDropdownOpen(false)
      }
      if (workloadDropdownRef.current && !workloadDropdownRef.current.contains(event.target as Node)) {
        setWorkloadDropdownOpen(false)
      }
      if (operationDropdownRef.current && !operationDropdownRef.current.contains(event.target as Node)) {
        setOperationDropdownOpen(false)
      }
    }

    document.addEventListener("mousedown", handleClickOutside)
    return () => {
      document.removeEventListener("mousedown", handleClickOutside)
    }
  }, [])

  const handleFileUpload = (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0]
    if (!file || !file.name.endsWith(".csv")) {
      alert("Only .csv files are supported.")
      return
    }
    setLoading(true)
    setFileName(file.name)
    Papa.parse(file, {
      header: true,
      skipEmptyLines: true,
      complete: (results) => {
        const data = results.data.map((entry: any) => {
          let auditData = {}
          try {
            auditData = JSON.parse(entry.AuditData || "{}")
          } catch (err) {
            console.warn("Failed to parse AuditData", err)
          }
          return {
            ...entry,
            FileName: auditData.ObjectId || "",
            Subject: auditData.Subject || "",
            MessageId: auditData.MessageId || auditData.InternetMessageId || "",
            TimeGenerated: entry.CreationDate,
            ClientIP: auditData.ClientIP || auditData.ClientIPAddress || "",
            CorrelationId: auditData.CorrelationId || auditData.CorrelationID || "",
            ModifiedProperties: auditData.ModifiedProperties
              ? JSON.stringify(auditData.ModifiedProperties, null, 2)
              : "N/A",
            Workload: auditData.Workload || entry.Workload || "Unknown",
            AuditDataRaw: entry.AuditData,
          }
        })
        setLogs(data)
        setLoading(false)
      },
    })
  }

  const userOptions = useMemo(() => Array.from(new Set(logs.map((e) => e.UserId || e.UserKey).filter(Boolean))), [logs])
  const workloadOptions = useMemo(() => Array.from(new Set(logs.map((e) => e.Workload).filter(Boolean))), [logs])
  const operationOptions = useMemo(() => Array.from(new Set(logs.map((e) => e.Operation).filter(Boolean))), [logs])

  const filteredLogs = useMemo(() => {
    return logs.filter((entry) => {
      // Multi-select filters - if no filters selected, show all
      const userMatch = userFilters.length === 0 || userFilters.includes(entry.UserId || entry.UserKey)
      const workloadMatch = workloadFilters.length === 0 || workloadFilters.includes(entry.Workload)
      const operationMatch = operationFilters.length === 0 || operationFilters.includes(entry.Operation)

      // Single-select filters
      const correlationMatch = correlationFilter ? entry.CorrelationId === correlationFilter : true
      const riskyMatch = showOnlyRisky ? riskyOps.includes(entry.Operation) : true

      // Search functionality
      const searchMatch = searchTerm ? JSON.stringify(entry).toLowerCase().includes(searchTerm.toLowerCase()) : true

      return userMatch && workloadMatch && operationMatch && correlationMatch && riskyMatch && searchMatch
    })
  }, [logs, userFilters, workloadFilters, operationFilters, correlationFilter, showOnlyRisky, searchTerm])

  const visibleLogs = filteredLogs.slice(0, visibleCount)

  const resetFilters = () => {
    setUserFilters([])
    setWorkloadFilters([])
    setOperationFilters([])
    setCorrelationFilter("")
    setShowOnlyRisky(false)
    setSearchTerm("")
    setVisibleCount(100)
  }

  // Toggle a filter in a multi-select array
  const toggleFilter = (type: "user" | "workload" | "operation", value: string) => {
    switch (type) {
      case "user":
        setUserFilters((prev) => (prev.includes(value) ? prev.filter((item) => item !== value) : [...prev, value]))
        break
      case "workload":
        setWorkloadFilters((prev) => (prev.includes(value) ? prev.filter((item) => item !== value) : [...prev, value]))
        break
      case "operation":
        setOperationFilters((prev) => (prev.includes(value) ? prev.filter((item) => item !== value) : [...prev, value]))
        break
    }
  }

  // Check if any filters are active
  const hasActiveFilters =
    userFilters.length > 0 ||
    workloadFilters.length > 0 ||
    operationFilters.length > 0 ||
    correlationFilter ||
    showOnlyRisky ||
    searchTerm

  const downloadNDJSON = () => {
    const ndjson = logs.map((l) => l.AuditDataRaw).join("\n")
    const blob = new Blob([ndjson], { type: "application/x-ndjson" })

    const url = URL.createObjectURL(blob)
    const a = document.createElement("a")
    a.href = url
    a.download = "ual-export.json"
    a.click()
    URL.revokeObjectURL(url)
  }

  const formatDate = (dateString: string) => {
    if (!dateString) return ""

    try {
      // Try different date parsing approaches
      let date

      // First try direct parsing
      date = new Date(dateString)

      // Check if the date is valid
      if (isNaN(date.getTime())) {
        // Try parsing ISO format with timezone handling
        if (dateString.includes("T") && dateString.includes("Z")) {
          // Already in ISO format, but might need special handling
          date = new Date(dateString)
        } else if (dateString.includes("/")) {
          // Handle MM/DD/YYYY format
          const parts = dateString.split("/")
          date = new Date(`${parts[2]}-${parts[0]}-${parts[1]}`)
        } else if (dateString.match(/^\d{4}-\d{2}-\d{2}/)) {
          // Already in YYYY-MM-DD format
          date = new Date(dateString)
        }
      }

      // Final check if date is valid
      if (isNaN(date.getTime())) {
        console.warn("Could not parse date:", dateString)
        return dateString // Return original string if we can't parse it
      }

      // Format the date
      return new Intl.DateTimeFormat("default", {
        year: "numeric",
        month: "short",
        day: "numeric",
        hour: "2-digit",
        minute: "2-digit",
        second: "2-digit",
      }).format(date)
    } catch (e) {
      console.warn("Error formatting date:", e)
      return dateString // Return original string on error
    }
  }

  return (
    <div className="min-h-screen bg-white dark:bg-gray-900 p-4 md:p-6">
      <div className="max-w-6xl mx-auto bg-white dark:bg-gray-900 shadow-xl rounded-xl border border-slate-200 dark:border-gray-700 overflow-hidden">
        {/* Header */}
        <div className="bg-gradient-to-r from-blue-600 to-indigo-600 p-6 text-white">
          <h1 className="text-3xl font-bold flex items-center gap-2">
            <Shield className="h-8 w-8" />
            UAL-Timeline-Builder (UTB)
          </h1>
          <p className="mt-2 opacity-90">Analyze and visualize Microsoft 365 Unified Audit Logs (or export to ndjson)</p>
          <p className="mt-2 text-sm font-medium text-yellow-200">
  ⚠️ This tool processes data entirely in your browser. No data is stored or transmitted to any server.
</p>
        </div>

        {/* Upload Section */}
        <div className="p-6 border-b border-slate-200 dark:border-gray-700">
          <div className="bg-slate-50 dark:bg-gray-800 rounded-xl p-6 border border-slate-200 dark:border-gray-700 transition-all hover:border-blue-300 dark:hover:border-blue-700">
            <div className="flex flex-col md:flex-row md:items-center gap-4">
              <div className="flex-1">
                <h2 className="text-lg font-semibold mb-1">Upload Audit Log</h2>
                <p className="text-slate-500 dark:text-gray-400 text-sm">Select a CSV file containing UAL data</p>
              </div>
              <div className="flex-shrink-0">
                <label className="inline-flex items-center gap-2 px-4 py-2 bg-blue-600 hover:bg-blue-700 text-white rounded-lg cursor-pointer transition-colors">
                  <input type="file" accept=".csv" onChange={handleFileUpload} className="hidden" />
                  <span>Select File</span>
                </label>
              </div>
            </div>
            {fileName && (
              <div className="mt-4 flex items-center gap-2 text-sm bg-blue-50 dark:bg-blue-900/30 text-blue-700 dark:text-blue-300 p-2 px-3 rounded-lg">
                <Info className="h-4 w-4" />
                <span>
                  Uploaded: <span className="font-medium">{fileName}</span>
                </span>
              </div>
            )}
          </div>
        </div>

        {loading ? (
          <div className="flex flex-col items-center justify-center py-32">
            <Loader2 className="h-12 w-12 text-blue-600 animate-spin mb-4" />
            <p className="text-slate-600 dark:text-gray-400 animate-pulse">Processing logs...</p>
          </div>
        ) : (
          <>
            {logs.length > 0 && (
              <div className="p-6 border-b border-slate-200 dark:border-gray-700">
                <div className="flex flex-col md:flex-row gap-4 items-start md:items-center justify-between mb-6">
                  <div className="flex-1">
                    <h2 className="text-xl font-semibold mb-1">Audit Log Timeline</h2>
                    <p className="text-slate-500 dark:text-gray-400 text-sm">{filteredLogs.length} events found</p>
                  </div>
                  <div className="flex flex-col sm:flex-row gap-2 w-full md:w-auto">
                    <div className="relative flex-1">
                      <Search className="absolute left-3 top-1/2 transform -translate-y-1/2 h-4 w-4 text-slate-400" />
                      <input
                        type="text"
                        placeholder="Search logs..."
                        value={searchTerm}
                        onChange={(e) => setSearchTerm(e.target.value)}
                        className="pl-9 pr-4 py-2 w-full rounded-lg border border-slate-300 dark:border-slate-700 bg-white dark:bg-gray-800 focus:ring-2 focus:ring-blue-500 focus:border-blue-500 dark:focus:ring-blue-600 dark:focus:border-blue-600 outline-none transition-all"
                      />
                      {searchTerm && (
                        <button
                          onClick={() => setSearchTerm("")}
                          className="absolute right-3 top-1/2 transform -translate-y-1/2 text-slate-400 hover:text-slate-600 dark:hover:text-gray-300"
                        >
                          <X className="h-4 w-4" />
                        </button>
                      )}
                    </div>
                    <button
                      onClick={downloadNDJSON}
                      className="flex items-center justify-center gap-2 px-4 py-2 bg-slate-800 hover:bg-slate-900 dark:bg-slate-700 dark:hover:bg-gray-600 text-white rounded-lg transition-colors"
                    >
                      <Download className="h-4 w-4" />
                      <span>Export to ndjson</span>
                    </button>
                  </div>
                </div>

                {/* Filters */}
                <div className="bg-slate-50 dark:bg-gray-800 rounded-xl p-4 mb-6">
                  <div className="flex items-center justify-between mb-4">
                    <h3 className="font-medium flex items-center gap-2">
                      <Filter className="h-4 w-4" />
                      Filters
                    </h3>
                    <button
                      onClick={resetFilters}
                      className={`text-sm flex items-center gap-1 transition-colors ${
                        hasActiveFilters
                          ? "text-blue-600 dark:text-blue-400 hover:text-blue-800 dark:hover:text-blue-300"
                          : "text-slate-400 dark:text-gray-400 cursor-not-allowed"
                      }`}
                      disabled={!hasActiveFilters}
                    >
                      <RefreshCw className="h-3 w-3" />
                      Reset All
                    </button>
                  </div>

                  <div className="grid grid-cols-1 md:grid-cols-3 gap-4">
                    {/* Multi-select User dropdown */}
                    <div ref={userDropdownRef} className="relative">
                      <label className="block text-sm font-medium text-slate-700 dark:text-gray-300 mb-1">Users</label>
                      <button
                        onClick={() => setUserDropdownOpen(!userDropdownOpen)}
                        className={`w-full flex items-center justify-between rounded-lg border px-3 py-2 text-left focus:outline-none focus:ring-2 focus:ring-blue-500 focus:border-blue-500 dark:focus:ring-blue-600 dark:focus:border-blue-600 transition-all ${
                          userFilters.length > 0
                            ? "border-blue-500 dark:border-blue-600 bg-blue-50 dark:bg-blue-900/20 text-blue-800 dark:text-blue-300"
                            : "border-slate-300 dark:border-slate-700 bg-white dark:bg-gray-700 text-slate-900 dark:text-gray-100"
                        }`}
                      >
                        <span className="truncate">
                          {userFilters.length === 0
                            ? "All Users"
                            : userFilters.length === 1
                              ? userFilters[0]
                              : `${userFilters.length} users selected`}
                        </span>
                        <ChevronDown
                          className={`h-4 w-4 transition-transform ${userDropdownOpen ? "rotate-180" : ""}`}
                        />
                      </button>

                      {userDropdownOpen && (
                        <div className="absolute z-10 mt-1 w-full bg-white dark:bg-gray-700 rounded-lg border border-slate-200 dark:border-slate-700 shadow-lg max-h-60 overflow-auto">
                          <div className="p-2 border-b border-slate-200 dark:border-slate-700 sticky top-0 bg-white dark:bg-gray-700 z-10">
                            <button
                              onClick={() => setUserFilters([])}
                              className="text-xs text-blue-600 dark:text-blue-400 hover:underline"
                            >
                              Clear all
                            </button>
                          </div>
                          <div className="p-1">
                            {userOptions.map((user) => (
                              <div
                                key={user}
                                className="flex items-center px-3 py-2 hover:bg-slate-100 dark:hover:bg-gray-600 rounded cursor-pointer"
                                onClick={() => toggleFilter("user", user)}
                              >
                                <div className="mr-2 h-4 w-4 rounded border flex items-center justify-center border-slate-300 dark:border-slate-600">
                                  {userFilters.includes(user) && (
                                    <Check className="h-3 w-3 text-blue-600 dark:text-blue-400" />
                                  )}
                                </div>
                                <span className="text-sm">{user}</span>
                              </div>
                            ))}
                          </div>
                        </div>
                      )}
                    </div>

                    {/* Multi-select Workload dropdown */}
                    <div ref={workloadDropdownRef} className="relative">
                      <label className="block text-sm font-medium text-slate-700 dark:text-gray-300 mb-1">
                        Workloads
                      </label>
                      <button
                        onClick={() => setWorkloadDropdownOpen(!workloadDropdownOpen)}
                        className={`w-full flex items-center justify-between rounded-lg border px-3 py-2 text-left focus:outline-none focus:ring-2 focus:ring-blue-500 focus:border-blue-500 dark:focus:ring-blue-600 dark:focus:border-blue-600 transition-all ${
                          workloadFilters.length > 0
                            ? "border-blue-500 dark:border-blue-600 bg-blue-50 dark:bg-blue-900/20 text-blue-800 dark:text-blue-300"
                            : "border-slate-300 dark:border-slate-700 bg-white dark:bg-gray-700 text-slate-900 dark:text-gray-100"
                        }`}
                      >
                        <span className="truncate">
                          {workloadFilters.length === 0
                            ? "All Workloads"
                            : workloadFilters.length === 1
                              ? workloadFilters[0]
                              : `${workloadFilters.length} workloads selected`}
                        </span>
                        <ChevronDown
                          className={`h-4 w-4 transition-transform ${workloadDropdownOpen ? "rotate-180" : ""}`}
                        />
                      </button>

                      {workloadDropdownOpen && (
                        <div className="absolute z-10 mt-1 w-full bg-white dark:bg-gray-700 rounded-lg border border-slate-200 dark:border-slate-700 shadow-lg max-h-60 overflow-auto">
                          <div className="p-2 border-b border-slate-200 dark:border-slate-700 sticky top-0 bg-white dark:bg-gray-700 z-10">
                            <button
                              onClick={() => setWorkloadFilters([])}
                              className="text-xs text-blue-600 dark:text-blue-400 hover:underline"
                            >
                              Clear all
                            </button>
                          </div>
                          <div className="p-1">
                            {workloadOptions.map((workload) => (
                              <div
                                key={workload}
                                className="flex items-center px-3 py-2 hover:bg-slate-100 dark:hover:bg-gray-600 rounded cursor-pointer"
                                onClick={() => toggleFilter("workload", workload)}
                              >
                                <div className="mr-2 h-4 w-4 rounded border flex items-center justify-center border-slate-300 dark:border-slate-600">
                                  {workloadFilters.includes(workload) && (
                                    <Check className="h-3 w-3 text-blue-600 dark:text-blue-400" />
                                  )}
                                </div>
                                <span className="text-sm">{workload}</span>
                              </div>
                            ))}
                          </div>
                        </div>
                      )}
                    </div>

                    {/* Multi-select Operation dropdown */}
                    <div ref={operationDropdownRef} className="relative">
                      <label className="block text-sm font-medium text-slate-700 dark:text-gray-300 mb-1">
                        Operations
                      </label>
                      <button
                        onClick={() => setOperationDropdownOpen(!operationDropdownOpen)}
                        className={`w-full flex items-center justify-between rounded-lg border px-3 py-2 text-left focus:outline-none focus:ring-2 focus:ring-blue-500 focus:border-blue-500 dark:focus:ring-blue-600 dark:focus:border-blue-600 transition-all ${
                          operationFilters.length > 0
                            ? "border-blue-500 dark:border-blue-600 bg-blue-50 dark:bg-blue-900/20 text-blue-800 dark:text-blue-300"
                            : "border-slate-300 dark:border-slate-700 bg-white dark:bg-gray-700 text-slate-900 dark:text-gray-100"
                        }`}
                      >
                        <span className="truncate">
                          {operationFilters.length === 0
                            ? "All Operations"
                            : operationFilters.length === 1
                              ? operationFilters[0]
                              : `${operationFilters.length} operations selected`}
                        </span>
                        <ChevronDown
                          className={`h-4 w-4 transition-transform ${operationDropdownOpen ? "rotate-180" : ""}`}
                        />
                      </button>

                      {operationDropdownOpen && (
                        <div className="absolute z-10 mt-1 w-full bg-white dark:bg-gray-700 rounded-lg border border-slate-200 dark:border-slate-700 shadow-lg max-h-60 overflow-auto">
                          <div className="p-2 border-b border-slate-200 dark:border-slate-700 sticky top-0 bg-white dark:bg-gray-700 z-10">
                            <button
                              onClick={() => setOperationFilters([])}
                              className="text-xs text-blue-600 dark:text-blue-400 hover:underline"
                            >
                              Clear all
                            </button>
                          </div>
                          <div className="p-1">
                            {operationOptions.map((operation) => (
                              <div
                                key={operation}
                                className="flex items-center px-3 py-2 hover:bg-slate-100 dark:hover:bg-gray-600 rounded cursor-pointer"
                                onClick={() => toggleFilter("operation", operation)}
                              >
                                <div className="mr-2 h-4 w-4 rounded border flex items-center justify-center border-slate-300 dark:border-slate-600">
                                  {operationFilters.includes(operation) && (
                                    <Check className="h-3 w-3 text-blue-600 dark:text-blue-400" />
                                  )}
                                </div>
                                <span className="text-sm">
                                  {operation}
                                  {riskyOps.includes(operation) && (
                                    <span className="ml-2 inline-flex items-center text-xs font-medium text-red-600 dark:text-red-400">
                                      <AlertTriangle className="h-3 w-3 mr-1" />
                                      Risky
                                    </span>
                                  )}
                                </span>
                              </div>
                            ))}
                          </div>
                        </div>
                      )}
                    </div>

                    {/* Risky Operations Toggle */}
                    <div className="md:col-span-3">
                      <label
                        className={`flex items-center gap-2 w-full p-2 rounded-lg cursor-pointer transition-colors ${
                          showOnlyRisky
                            ? "bg-blue-50 dark:bg-blue-900/20 text-blue-800 dark:text-blue-300"
                            : "hover:bg-slate-100 dark:hover:bg-gray-700"
                        }`}
                      >
                        <div className="relative flex items-center">
                          <input
                            type="checkbox"
                            checked={showOnlyRisky}
                            onChange={(e) => setShowOnlyRisky(e.target.checked)}
                            className="peer sr-only"
                          />
                          <div
                            className={`h-5 w-5 rounded border transition-colors ${
                              showOnlyRisky
                                ? "bg-blue-600 border-blue-600 dark:bg-blue-600 dark:border-blue-600"
                                : "border-slate-300 dark:border-slate-600"
                            }`}
                          ></div>
                          <svg
                            className="absolute h-3 w-3 text-white left-1 top-1 opacity-0 peer-checked:opacity-100 transition-opacity"
                            fill="none"
                            viewBox="0 0 24 24"
                            stroke="currentColor"
                            strokeWidth={3}
                          >
                            <path strokeLinecap="round" strokeLinejoin="round" d="M5 13l4 4L19 7" />
                          </svg>
                        </div>
                        <span className="text-sm font-medium">Show only risky operations</span>
                      </label>
                    </div>
                  </div>

                  {/* Active Filter Chips */}
                  {hasActiveFilters && (
                    <div className="mt-4 flex flex-wrap gap-2">
                      {userFilters.map((user) => (
                        <div
                          key={`user-${user}`}
                          className="inline-flex items-center gap-2 px-3 py-1 bg-blue-100 dark:bg-blue-900/40 text-blue-800 dark:text-blue-300 rounded-full text-sm"
                        >
                          <span>User: {user}</span>
                          <button
                            onClick={() => toggleFilter("user", user)}
                            className="hover:bg-blue-200 dark:hover:bg-blue-800 rounded-full p-0.5"
                          >
                            <X className="h-3 w-3" />
                            <span className="sr-only">Remove user filter</span>
                          </button>
                        </div>
                      ))}

                      {workloadFilters.map((workload) => (
                        <div
                          key={`workload-${workload}`}
                          className="inline-flex items-center gap-2 px-3 py-1 bg-blue-100 dark:bg-blue-900/40 text-blue-800 dark:text-blue-300 rounded-full text-sm"
                        >
                          <span>Workload: {workload}</span>
                          <button
                            onClick={() => toggleFilter("workload", workload)}
                            className="hover:bg-blue-200 dark:hover:bg-blue-800 rounded-full p-0.5"
                          >
                            <X className="h-3 w-3" />
                            <span className="sr-only">Remove workload filter</span>
                          </button>
                        </div>
                      ))}

                      {operationFilters.map((operation) => (
                        <div
                          key={`operation-${operation}`}
                          className="inline-flex items-center gap-2 px-3 py-1 bg-blue-100 dark:bg-blue-900/40 text-blue-800 dark:text-blue-300 rounded-full text-sm"
                        >
                          <span>Operation: {operation}</span>
                          <button
                            onClick={() => toggleFilter("operation", operation)}
                            className="hover:bg-blue-200 dark:hover:bg-blue-800 rounded-full p-0.5"
                          >
                            <X className="h-3 w-3" />
                            <span className="sr-only">Remove operation filter</span>
                          </button>
                        </div>
                      ))}

                      {correlationFilter && (
                        <div className="inline-flex items-center gap-2 px-3 py-1 bg-blue-100 dark:bg-blue-900/40 text-blue-800 dark:text-blue-300 rounded-full text-sm">
                          <span>Correlation ID: {correlationFilter.substring(0, 10)}...</span>
                          <button
                            onClick={() => setCorrelationFilter("")}
                            className="hover:bg-blue-200 dark:hover:bg-blue-800 rounded-full p-0.5"
                          >
                            <X className="h-3 w-3" />
                            <span className="sr-only">Remove correlation ID filter</span>
                          </button>
                        </div>
                      )}

                      {showOnlyRisky && (
                        <div className="inline-flex items-center gap-2 px-3 py-1 bg-blue-100 dark:bg-blue-900/40 text-blue-800 dark:text-blue-300 rounded-full text-sm">
                          <span>Risky Operations Only</span>
                          <button
                            onClick={() => setShowOnlyRisky(false)}
                            className="hover:bg-blue-200 dark:hover:bg-blue-800 rounded-full p-0.5"
                          >
                            <X className="h-3 w-3" />
                            <span className="sr-only">Remove risky operations filter</span>
                          </button>
                        </div>
                      )}

                      {searchTerm && (
                        <div className="inline-flex items-center gap-2 px-3 py-1 bg-blue-100 dark:bg-blue-900/40 text-blue-800 dark:text-blue-300 rounded-full text-sm">
                          <span>Search: {searchTerm}</span>
                          <button
                            onClick={() => setSearchTerm("")}
                            className="hover:bg-blue-200 dark:hover:bg-blue-800 rounded-full p-0.5"
                          >
                            <X className="h-3 w-3" />
                            <span className="sr-only">Remove search filter</span>
                          </button>
                        </div>
                      )}
                    </div>
                  )}
                </div>

                {/* Timeline Cards */}
                <div className="space-y-4">
                  {visibleLogs.length === 0 ? (
                    <div className="text-center py-12 bg-slate-50 dark:bg-gray-800 rounded-xl">
                      <p className="text-slate-500 dark:text-gray-400">No logs match your current filters</p>
                      <button onClick={resetFilters} className="mt-2 text-blue-600 dark:text-blue-400 hover:underline">
                        Reset filters
                      </button>
                    </div>
                  ) : (
                    visibleLogs.map((entry, i) => {
                      const isRisky = riskyOps.includes(entry.Operation)
                      return (
                        <div
                          key={i}
                          className={`
                            border rounded-xl overflow-hidden shadow-sm transition-all hover:shadow-md
                            ${
                              isRisky
                                ? "border-red-200 dark:border-red-900 bg-gradient-to-r from-red-50 to-white dark:from-red-950/30 dark:to-gray-900"
                                : "border-slate-200 dark:border-gray-700 bg-white dark:bg-gray-900"
                            }
                          `}
                        >
                          <div className="p-4">
                            <div className="flex flex-col md:flex-row md:items-center gap-3 mb-3">
                              <div className="flex items-center gap-2 text-sm text-slate-500 dark:text-gray-400">
                                <Clock className="h-4 w-4" />
                                <span>{formatDate(entry.TimeGenerated)}</span>
                              </div>

                              {isRisky && (
                                <div className="flex items-center gap-1 text-xs font-medium text-red-600 dark:text-red-400 bg-red-50 dark:bg-red-950/50 px-2 py-1 rounded-full">
                                  <AlertTriangle className="h-3 w-3" />
                                  <span>Risky Operation</span>
                                </div>
                              )}
                            </div>

                            <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-3 gap-3 mb-3">
                              <div className="flex items-center gap-2">
                                <User className="h-4 w-4 text-slate-400" />
                                <div className="text-sm">
                                  <div className="font-medium">{entry.UserId || entry.UserKey}</div>
                                  <div className="text-xs text-slate-500 dark:text-gray-400">User</div>
                                </div>
                              </div>

                              <div className="flex items-center gap-2">
                                <Shield className="h-4 w-4 text-slate-400" />
                                <div className="text-sm">
                                  <div className="font-medium">{entry.Operation}</div>
                                  <div className="text-xs text-slate-500 dark:text-gray-400">Operation</div>
                                </div>
                              </div>

                              <div className="flex items-center gap-2">
                                <Globe className="h-4 w-4 text-slate-400" />
                                <div className="text-sm">
                                  <div className="font-medium">{entry.Workload}</div>
                                  <div className="text-xs text-slate-500 dark:text-gray-400">Workload</div>
                                </div>
                              </div>

                              <div className="flex items-center gap-2">
                                <MapPin className="h-4 w-4 text-slate-400" />
                                <div className="text-sm">
                                  <div className="font-medium">{entry.ClientIP || "N/A"}</div>
                                  <div className="text-xs text-slate-500 dark:text-gray-400">Client IP</div>
                                </div>
                              </div>

                              <div className="flex items-center gap-2 col-span-1 md:col-span-2">
                                <Hash className="h-4 w-4 text-slate-400" />
                                <div className="text-sm">
                                  <div className="font-medium flex items-center gap-1">
                                    <span className="truncate max-w-[200px]">{entry.CorrelationId || "N/A"}</span>
                                    {entry.CorrelationId && (
                                      <button
                                        onClick={() => setCorrelationFilter(entry.CorrelationId)}
                                        className="text-xs text-blue-600 dark:text-blue-400 hover:underline"
                                      >
                                        Filter
                                      </button>
                                    )}
                                  </div>
                                  <div className="text-xs text-slate-500 dark:text-gray-400">Correlation ID</div>
                                </div>
                              </div>
                            </div>

                            <details className="group">
                              <summary className="cursor-pointer list-none flex items-center gap-1 text-sm font-medium text-slate-600 dark:text-gray-300 hover:text-blue-600 dark:hover:text-blue-400 transition-colors">
                                <ChevronDown className="h-4 w-4 transition-transform group-open:rotate-180" />
                                View Details
                              </summary>
                              <div className="mt-3 pt-3 border-t border-slate-200 dark:border-slate-800 text-sm space-y-3">
                                {entry.FileName && (
                                  <div>
                                    <div className="font-medium mb-1">File</div>
                                    <div className="text-slate-700 dark:text-gray-300">{entry.FileName}</div>
                                  </div>
                                )}

                                {entry.Subject && (
                                  <div>
                                    <div className="font-medium mb-1">Subject</div>
                                    <div className="text-slate-700 dark:text-gray-300">{entry.Subject}</div>
                                  </div>
                                )}

                                {entry.MessageId && (
                                  <div>
                                    <div className="font-medium mb-1">Message ID</div>
                                    <div className="text-slate-700 dark:text-gray-300 break-all">{entry.MessageId}</div>
                                  </div>
                                )}

                                <div>
                                  <div className="font-medium mb-1">Modified Properties</div>
                                  <pre className="bg-slate-50 dark:bg-gray-800 p-3 rounded-lg overflow-auto text-xs whitespace-pre-wrap">
                                    {entry.ModifiedProperties}
                                  </pre>
                                </div>

                                <div>
                                  <div className="font-medium mb-1">Raw Event</div>
                                  <div className="relative">
                                    <pre className="bg-slate-50 dark:bg-gray-800 p-3 rounded-lg overflow-auto text-xs max-h-[300px]">
                                      {JSON.stringify(entry, null, 2)}
                                    </pre>
                                    <button
                                      className="absolute top-2 right-2 bg-slate-200 dark:bg-slate-700 hover:bg-slate-300 dark:hover:bg-gray-500 text-slate-700 dark:text-gray-300 p-1 rounded-md transition-colors"
                                      onClick={() => {
                                        navigator.clipboard.writeText(JSON.stringify(entry, null, 2))
                                      }}
                                      title="Copy to clipboard"
                                    >
                                      <svg
                                        xmlns="http://www.w3.org/2000/svg"
                                        className="h-4 w-4"
                                        fill="none"
                                        viewBox="0 0 24 24"
                                        stroke="currentColor"
                                      >
                                        <path
                                          strokeLinecap="round"
                                          strokeLinejoin="round"
                                          strokeWidth={2}
                                          d="M8 5H6a2 2 0 00-2 2v12a2 2 0 002 2h10a2 2 0 002-2v-1M8 5a2 2 0 002 2h2a2 2 0 002-2M8 5a2 2 0 012-2h2a2 2 0 012 2m0 0h2a2 2 0 012 2v3m2 4H10m0 0l3-3m-3 3l3 3"
                                        />
                                      </svg>
                                    </button>
                                  </div>
                                </div>
                              </div>
                            </details>
                          </div>
                        </div>
                      )
                    })
                  )}
                </div>

                {visibleCount < filteredLogs.length && (
                  <div className="text-center mt-8">
                    <button
                      onClick={() => setVisibleCount(visibleCount + 100)}
                      className="inline-flex items-center gap-2 px-6 py-3 bg-slate-100 hover:bg-slate-200 dark:bg-slate-700 dark:hover:bg-gray-600 text-slate-700 dark:text-gray-300 rounded-lg transition-colors"
                    >
                      <span>Load More</span>
                      <span className="text-sm text-slate-500 dark:text-gray-400">
                        ({visibleCount} / {filteredLogs.length})
                      </span>
                    </button>
                  </div>
                )}
              </div>
            )}

            {logs.length === 0 && !loading && (
              <div className="flex flex-col items-center justify-center py-20">
                <Calendar className="h-16 w-16 text-slate-300 dark:text-slate-700 mb-4" />
                <h3 className="text-xl font-medium text-slate-700 dark:text-gray-300 mb-2">No Audit Logs</h3>
                <p className="text-slate-500 dark:text-gray-400 text-center max-w-md">
                  Upload a CSV file containing Microsoft 365 Unified Audit Logs to get started
                </p>
              </div>
            )}
          </>
        )}
      </div>

      <div className="mt-4 text-center text-xs text-slate-500 dark:text-gray-400">
      <a href="https://github.com/SagaLabs/UAL-Timeline-Builder">https://github.com/SagaLabs/UAL-Timeline-Builder</a>
      </div>
    </div>
  )
}

