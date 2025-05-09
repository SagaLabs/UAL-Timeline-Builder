"use client"

import type React from "react"

import { useState, useMemo, useRef, useEffect, useCallback } from "react"
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
  Mail,
  Moon,
  RefreshCw,
  Search,
  Shield,
  Sun,
  User,
  X,
  Inbox,
  Monitor,
  Laptop,
  Network,
  Code,
  Trash2,
  Save,
  Upload,
  PencilIcon,
  FileText,
  Map,
  ChartBar
} from "lucide-react"

interface RuleDetails {
  Name: string;
  Actions: string[];
  Conditions: string[];
  Enabled: boolean;
  Priority: string;
  StopProcessingRules: boolean;
}

interface UserAgentInfo {
  browser: string;
  browserVersion: string;
  os: string;
  osVersion: string;
  device: string;
  isMobile: boolean;
  raw: string;
}

interface AuditData {
  ObjectId?: string;
  Subject?: string;
  MessageId?: string;
  InternetMessageId?: string;
  ClientIP?: string;
  ClientIPAddress?: string;
  CorrelationId?: string;
  CorrelationID?: string;
  ModifiedProperties?: any[];
  Parameters?: any[];
  Workload?: string;
  UserAgent?: string;
  MailboxOwnerUPN?: string;
  MailAccessType?: string;
  ClientInfoString?: string;
  AppAccessContext?: {
    ClientAppId?: string;
    ClientIPAddress?: string;
  };
  OperationProperties?: Array<{
    Name: string;
    Value: string;
  }>;
  AddOnName?: string;
  AppDistributionMode?: string;
  AzureADAppId?: string;
  ResourceSpecificApplicationPermissions?: string[];
  ChatThreadId?: string;
  DeviceId?: string;
  ExtendedProperties?: any[];
}

interface LogEntry {
  AuditData: string | AuditData;
  Operation: string;
  CreationDate: string;
  Workload?: string;
  UserId?: string;
  UserKey?: string;
  RuleDetails?: RuleDetails;
  FileName?: string;
  Subject?: string;
  MessageId?: string;
  TimeGenerated?: string;
  ClientIP?: string;
  ClientIPAddress?: string;
  CorrelationId?: string;
  ModifiedProperties?: string;
  AuditDataRaw?: string;
  UserAgent?: string;
  UserAgentInfo?: UserAgentInfo;
}

// Add TimelineEntry interface
interface TimelineEntry {
  id: string;
  timestamp: string;
  title: string;
  description: string;
  type: 'info' | 'warning' | 'error';
  logEntry: LogEntry;
  note?: string;
}

interface ModifiedProperty {
  Name: string;
  Value: string;
}

interface Parameter {
  Name: string;
  Value: string;
}

interface ChartData {
  labels: string[];
  datasets: {
    label: string;
    data: number[];
    borderColor?: string;
    backgroundColor?: string;
    fill?: boolean;
  }[];
}

// Add debounce utility function
const debounce = <T extends (...args: any[]) => any>(
  func: T,
  wait: number
): ((...args: Parameters<T>) => void) => {
  let timeout: NodeJS.Timeout;
  return (...args: Parameters<T>) => {
    clearTimeout(timeout);
    timeout = setTimeout(() => func(...args), wait);
  };
};

// Add useDarkMode hook
const useDarkMode = () => {
  const [darkMode, setDarkMode] = useState(false);
  const [mounted, setMounted] = useState(false);

  useEffect(() => {
    // Get initial dark mode preference
    const savedPreference = localStorage.getItem('darkMode');
    const systemPreference = window.matchMedia('(prefers-color-scheme: dark)').matches;
    
    setDarkMode(savedPreference !== null ? JSON.parse(savedPreference) : systemPreference);
    setMounted(true);
  }, []);

  useEffect(() => {
    if (mounted) {
      if (darkMode) {
        document.documentElement.classList.add('dark');
        localStorage.setItem('darkMode', 'true');
      } else {
        document.documentElement.classList.remove('dark');
        localStorage.setItem('darkMode', 'false');
      }
    }
  }, [darkMode, mounted]);

  // Listen for system theme changes
  useEffect(() => {
    if (mounted) {
      const mediaQuery = window.matchMedia('(prefers-color-scheme: dark)');
      const handleChange = (e: MediaQueryListEvent) => {
        const savedPreference = localStorage.getItem('darkMode');
        if (savedPreference === null) {
          setDarkMode(e.matches);
        }
      };

      mediaQuery.addEventListener('change', handleChange);
      return () => mediaQuery.removeEventListener('change', handleChange);
    }
  }, [mounted]);

  return [darkMode, setDarkMode, mounted] as const;
};

interface AuthenticationStats {
  ip: string;
  userId: string;
  count: number;
  firstSeen: string;
  lastSeen: string;
  operations: Set<string>;
}

interface ReportSection {
  title: string;
  content: string;
}

interface CaseInfo {
  caseName: string;
  caseNumber: string;
  analystName: string;
  analysisDate: string;
  company: string;
}

export default function UALTimelineBuilder() {
  const [logs, setLogs] = useState<LogEntry[]>([])
  const [fileNames, setFileNames] = useState<string[]>([])
  const [userFilters, setUserFilters] = useState<string[]>([])
  const [workloadFilters, setWorkloadFilters] = useState<string[]>([])
  const [operationFilters, setOperationFilters] = useState<string[]>([])
  const [ipFilters, setIpFilters] = useState<string[]>([])
  const [correlationFilter, setCorrelationFilter] = useState("")
  const [showOnlyRisky, setShowOnlyRisky] = useState(false)
  const [visibleCount, setVisibleCount] = useState(100)
  const [loading, setLoading] = useState(false)
  const [searchTerm, setSearchTerm] = useState("")
  const [darkMode, setDarkMode, mounted] = useDarkMode()
  const [showToast, setShowToast] = useState(false)
  const [toastMessage, setToastMessage] = useState("")
  const [showCaseInfoModal, setShowCaseInfoModal] = useState(false)
  const [caseInfo, setCaseInfo] = useState<CaseInfo>({
    caseName: "",
    caseNumber: "",
    analystName: "",
    analysisDate: new Date().toISOString().split('T')[0],
    company: ""
  })

  const [userDropdownOpen, setUserDropdownOpen] = useState(false)
  const [workloadDropdownOpen, setWorkloadDropdownOpen] = useState(false)
  const [operationDropdownOpen, setOperationDropdownOpen] = useState(false)
  const [ipDropdownOpen, setIpDropdownOpen] = useState(false)
  const [userSearchTerm, setUserSearchTerm] = useState("")
  const [workloadSearchTerm, setWorkloadSearchTerm] = useState("")
  const [operationSearchTerm, setOperationSearchTerm] = useState("")
  const [ipSearchTerm, setIpSearchTerm] = useState("")
  const [showIPExportModal, setShowIPExportModal] = useState(false)
  const [timelineEvents, setTimelineEvents] = useState<TimelineEntry[]>([]);
  const [showTimeline, setShowTimeline] = useState(false);
  const [activeNoteId, setActiveNoteId] = useState<string | null>(null);
  const [timelineSortAsc, setTimelineSortAsc] = useState(true); // Add state for sort direction
  const [showAuthBaseline, setShowAuthBaseline] = useState(false);
  const [authBaselineData, setAuthBaselineData] = useState<AuthenticationStats[]>([]);
  const [reportSections, setReportSections] = useState<ReportSection[]>([]);
  const [isDragging, setIsDragging] = useState(false)

  const userDropdownRef = useRef<HTMLDivElement>(null)
  const workloadDropdownRef = useRef<HTMLDivElement>(null)
  const operationDropdownRef = useRef<HTMLDivElement>(null)
  const ipDropdownRef = useRef<HTMLDivElement>(null)

  const riskyOps = [
    "UpdateInboxRule",
    "New-InboxRule",
    "Remove-InboxRule",
    "Add-MailboxPermission",
    "Set-Mailbox",
    "Set-MailboxAutoReplyConfiguration",
    "Add member to role.",
    "Add user.",
    "Add delegated permission grant.",
    "Set-AdminAuditLogConfig",
    "Update application – Certificates and secrets management ",
    "Consent to application.",
    "Add service principal.",
    "Update application.",
    "Add application.",
    "Add application permission.",
    "Update PasswordProfile.",
    "Change user password.",
    "Add owner to application."
  ]

  // Add list of email operations
  const emailOps = [
    "MailItemsAccessed",
    "Send",
    "UpdateInboxRule",
    "New-InboxRule",
    "Set-Mailbox",
    "Set-MailboxAutoReplyConfiguration",
    "Add-MailboxPermission",
    "Remove-MailboxPermission",
    "Update-MailboxPermission",
    "Move-Mailbox",
    "New-Mailbox",
    "Remove-Mailbox",
    "Set-MailboxRegionalConfiguration",
    "Set-MailboxCalendarConfiguration",
    "Set-MailboxMessageConfiguration"
  ]

  // Add debounced search handler
  const debouncedSetSearchTerm = useMemo(
    () => debounce((value: string) => setSearchTerm(value), 300),
    []
  );

  // Optimize search fields
  const searchableFields = useMemo(() => {
    return logs.map(entry => ({
      id: entry.UserId || entry.UserKey || '',
      operation: entry.Operation || '',
      workload: entry.Workload || '',
      subject: entry.Subject || '',
      messageId: entry.MessageId || '',
      correlationId: entry.CorrelationId || '',
      clientIP: entry.ClientIP || '',
      fileName: entry.FileName || '',
      userAgent: entry.UserAgent || ''
    }));
  }, [logs]);

  // Update the dark mode button render logic
  const renderDarkModeButton = () => {
    if (!mounted) {
      // Return a placeholder with the same dimensions to avoid layout shift
      return (
        <button
          className="p-2 rounded-lg bg-white/10 hover:bg-white/20 transition-colors"
          aria-hidden="true"
        >
          <div className="h-6 w-6" />
        </button>
      );
    }

    return (
      <button
        onClick={() => setDarkMode(!darkMode)}
        className="p-2 rounded-lg bg-white/10 hover:bg-white/20 transition-colors"
        title={darkMode ? "Switch to light mode" : "Switch to dark mode"}
      >
        {darkMode ? <Sun className="h-6 w-6" /> : <Moon className="h-6 w-6" />}
      </button>
    );
  };

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
      if (ipDropdownRef.current && !ipDropdownRef.current.contains(event.target as Node)) {
        setIpDropdownOpen(false)
      }
    }

    document.addEventListener("mousedown", handleClickOutside)
    return () => {
      document.removeEventListener("mousedown", handleClickOutside)
    }
  }, [])

  const parseUserAgent = (userAgent: string): UserAgentInfo => {
    const info: UserAgentInfo = {
      browser: "Unknown",
      browserVersion: "Unknown",
      os: "Unknown",
      osVersion: "Unknown",
      device: "Unknown",
      isMobile: false,
      raw: userAgent
    };

    try {
      // Basic browser detection
      if (userAgent.includes("Firefox")) {
        info.browser = "Firefox";
        const match = userAgent.match(/Firefox\/(\d+\.\d+)/);
        if (match) info.browserVersion = match[1];
      } else if (userAgent.includes("Chrome")) {
        info.browser = "Chrome";
        const match = userAgent.match(/Chrome\/(\d+\.\d+)/);
        if (match) info.browserVersion = match[1];
      } else if (userAgent.includes("Safari")) {
        info.browser = "Safari";
        const match = userAgent.match(/Version\/(\d+\.\d+)/);
        if (match) info.browserVersion = match[1];
      } else if (userAgent.includes("Edge")) {
        info.browser = "Edge";
        const match = userAgent.match(/Edge\/(\d+\.\d+)/);
        if (match) info.browserVersion = match[1];
      } else if (userAgent.includes("MSIE") || userAgent.includes("Trident/")) {
        info.browser = "Internet Explorer";
        const match = userAgent.match(/MSIE (\d+\.\d+)/) || userAgent.match(/rv:(\d+\.\d+)/);
        if (match) info.browserVersion = match[1];
      }

      // OS detection
      if (userAgent.includes("Windows")) {
        info.os = "Windows";
        const match = userAgent.match(/Windows NT (\d+\.\d+)/);
        if (match) {
          const version = parseFloat(match[1]);
          info.osVersion = version === 10 ? "10" : version === 6.3 ? "8.1" : version === 6.2 ? "8" : match[1];
        }
      } else if (userAgent.includes("Macintosh")) {
        info.os = "macOS";
        const match = userAgent.match(/Mac OS X (\d+[._]\d+)/);
        if (match) info.osVersion = match[1].replace(/_/g, ".");
      } else if (userAgent.includes("Linux")) {
        info.os = "Linux";
        const match = userAgent.match(/Linux ([^;)]+)/);
        if (match) info.osVersion = match[1];
      } else if (userAgent.includes("Android")) {
        info.os = "Android";
        const match = userAgent.match(/Android (\d+\.\d+)/);
        if (match) info.osVersion = match[1];
      } else if (userAgent.includes("iOS")) {
        info.os = "iOS";
        const match = userAgent.match(/OS (\d+_\d+)/);
        if (match) info.osVersion = match[1].replace(/_/g, ".");
      }

      // Device detection
      if (userAgent.includes("Mobile")) {
        info.device = "Mobile";
        info.isMobile = true;
      } else if (userAgent.includes("Tablet")) {
        info.device = "Tablet";
        info.isMobile = true;
      } else {
        info.device = "Desktop";
        info.isMobile = false;
      }
    } catch (e) {
      console.warn("Error parsing user agent:", e);
    }

    return info;
  };

  const handleFileUpload = (e: React.ChangeEvent<HTMLInputElement>) => {
    const files = Array.from(e.target.files || [])
    if (files.length === 0) return

    // Check if all files are CSV
    if (!files.every(file => file.name.endsWith(".csv"))) {
      alert("Only .csv files are supported.")
      return
    }

    setLoading(true)
    setFileNames(files.map(f => f.name))

    // Process files sequentially
    let allLogs: LogEntry[] = []
    let processedCount = 0

    const processFile = (file: File) => {
      Papa.parse(file, {
        header: true,
        skipEmptyLines: true,
        complete: (results) => {
          const data = results.data.map((entry: LogEntry) => {
            let auditData: AuditData = {}
            try {
              auditData = typeof entry.AuditData === 'string' ? JSON.parse(entry.AuditData) : entry.AuditData
            } catch (err) {
              console.warn("Failed to parse AuditData", err)
            }
            let ruleDetails = null
            if (entry.Operation === "UpdateInboxRule" || entry.Operation === "New-InboxRule" || entry.Operation === "Remove-InboxRule") {
              try {
                if (auditData.ModifiedProperties) {
                  ruleDetails = extractRuleDetails(auditData.ModifiedProperties)
                } else if (auditData.Parameters) {
                  ruleDetails = extractRuleDetailsFromParameters(auditData.Parameters)
                }
              } catch (err) {
                console.warn("Failed to parse rule details", err)
              }
            }

            // Parse User Agent if available
            let userAgentInfo: UserAgentInfo | undefined;
            if (auditData.UserAgent) {
              userAgentInfo = parseUserAgent(auditData.UserAgent);
            }

            return {
              ...entry,
              FileName: auditData.ObjectId ?? "",
              Subject: auditData.Subject ?? "",
              MessageId: auditData.MessageId ?? auditData.InternetMessageId ?? "",
              TimeGenerated: entry.CreationDate,
              ClientIP: auditData.ClientIP ?? auditData.ClientIPAddress ?? "",
              CorrelationId: auditData.CorrelationId ?? auditData.CorrelationID ?? "",
              ModifiedProperties: auditData.ModifiedProperties
                ? JSON.stringify(auditData.ModifiedProperties, null, 2)
                : "N/A",
              Workload: auditData.Workload ?? entry.Workload ?? "Unknown",
              AuditDataRaw: entry.AuditData,
              RuleDetails: ruleDetails ?? undefined,
              UserAgent: auditData.UserAgent,
              UserAgentInfo: userAgentInfo
            }
          })
          allLogs = [...allLogs, ...data]
          processedCount++

          // If all files are processed, update the state
          if (processedCount === files.length) {
            setLogs(allLogs)
            setLoading(false)
          }
        },
        error: (error) => {
          console.error("Error parsing file:", error)
          processedCount++
          if (processedCount === files.length) {
            setLogs(allLogs)
            setLoading(false)
          }
        }
      })
    }

    // Process each file
    files.forEach(processFile)
  }

  const extractRuleDetails = (modifiedProperties: ModifiedProperty[]): RuleDetails | null => {
    if (!Array.isArray(modifiedProperties)) return null;

    const details: RuleDetails = {
    Name: "",
    Actions: [],
    Conditions: [],
    Enabled: true,
    Priority: "",
    StopProcessingRules: false,
    };

    modifiedProperties.forEach((prop) => {
      if (prop.Name === "Name" && prop.Value) {
        details.Name = prop.Value;
      } else if (prop.Name === "ForwardTo" || prop.Name === "RedirectTo") {
        details.Actions.push(`Forward to: ${prop.Value}`);
      } else if (prop.Name === "DeleteMessage") {
        details.Actions.push("Delete message");
      } else if (prop.Name === "MoveToFolder") {
        details.Actions.push(`Move to folder: ${prop.Value}`);
      } else if (prop.Name === "From") {
        details.Conditions.push(`From: ${prop.Value}`);
      } else if (prop.Name === "SubjectContainsWords") {
        details.Conditions.push(`Subject contains: ${prop.Value}`);
      } else if (prop.Name === "BodyContainsWords") {
        details.Conditions.push(`Body contains: ${prop.Value}`);
      } else if (prop.Name === "Enabled") {
        details.Enabled = prop.Value === "True";
      } else if (prop.Name === "Priority") {
        details.Priority = prop.Value;
      } else if (prop.Name === "StopProcessingRules") {
        details.StopProcessingRules = prop.Value === "True";
      }
    });

    return details;
  };

  const extractRuleDetailsFromParameters = (parameters: Parameter[]): RuleDetails | null => {
    if (!Array.isArray(parameters)) return null;

    const details: RuleDetails = {
      Name: "",
      Actions: [],
      Conditions: [],
      Enabled: true,
      Priority: "",
      StopProcessingRules: false,
    };

  parameters.forEach((param) => {
    if (param.Name === "Name" && param.Value) {
      details.Name = param.Value
    } else if (
      param.Name === "ForwardTo" ||
      param.Name === "ForwardAsAttachmentTo" ||
      (param.Name === "RedirectTo" && param.Value)
    ) {
      details.Actions.push(`Forward to: ${param.Value.replace(/\[|\]/g, "")}`)
    } else if (param.Name === "DeleteMessage" && param.Value === "True") {
      details.Actions.push("Delete message")
    } else if (param.Name === "MoveToFolder" && param.Value) {
      details.Actions.push(`Move to folder: ${param.Value}`)
    } else if (param.Name === "CopyToFolder" && param.Value) {
      details.Actions.push(`Copy to folder: ${param.Value}`)
    } else if (param.Name === "From" && param.Value) {
      details.Conditions.push(`From: ${param.Value.replace(/\[|\]/g, "")}`)
    } else if (param.Name === "SubjectContainsWords" && param.Value) {
      details.Conditions.push(`Subject contains: ${param.Value}`)
    } else if (param.Name === "BodyContainsWords" && param.Value) {
      details.Conditions.push(`Body contains: ${param.Value}`)
    } else if (param.Name === "Enabled" && param.Value) {
      details.Enabled = param.Value === "True"
    } else if (param.Name === "Priority" && param.Value) {
      details.Priority = param.Value
    } else if (param.Name === "StopProcessingRules" && param.Value) {
      details.StopProcessingRules = param.Value === "True"
    }
  })

  return details
}

  const userOptions = useMemo(() => Array.from(new Set(logs.map((e) => e.UserId || e.UserKey).filter(Boolean))), [logs])
  const workloadOptions = useMemo(() => Array.from(new Set(logs.map((e) => e.Workload).filter(Boolean))), [logs])
  const operationOptions = useMemo(() => Array.from(new Set(logs.map((e) => e.Operation).filter(Boolean))), [logs])

  // Filter options based on search terms
  const filteredUserOptions = useMemo(() => {
    if (!userSearchTerm) return userOptions;
    return userOptions.filter(user => 
      user.toLowerCase().includes(userSearchTerm.toLowerCase())
    );
  }, [userOptions, userSearchTerm]);

  const filteredWorkloadOptions = useMemo(() => {
    if (!workloadSearchTerm) return workloadOptions;
    return workloadOptions.filter(workload => 
      workload.toLowerCase().includes(workloadSearchTerm.toLowerCase())
    );
  }, [workloadOptions, workloadSearchTerm]);

  const filteredOperationOptions = useMemo(() => {
    if (!operationSearchTerm) return operationOptions;
    return operationOptions.filter(operation => 
      operation.toLowerCase().includes(operationSearchTerm.toLowerCase())
    );
  }, [operationOptions, operationSearchTerm]);

  const ipOptions = useMemo(() => Array.from(new Set(logs.map((e) => e.ClientIP || e.ClientIPAddress).filter(Boolean))), [logs])

  const filteredIpOptions = useMemo(() => {
    if (!ipSearchTerm) return ipOptions;
    return ipOptions.filter(ip => 
      ip.toLowerCase().includes(ipSearchTerm.toLowerCase())
    );
  }, [ipOptions, ipSearchTerm]);

  // Add a helper function for consistent timestamp parsing
  const parseTimestamp = (timestamp: string) => {
    try {
      // Handle both formats and ensure proper timezone handling
      const date = new Date(timestamp);
      if (isNaN(date.getTime())) {
        // If direct parsing fails, try to parse the format: YYYY-MM-DDThh:mm:ss.000Z
        const [datePart, timePart] = timestamp.split('T');
        const [year, month, day] = datePart.split('-').map(Number);
        const [hours, minutes, seconds] = timePart.split(':').map(n => parseInt(n));
        return new Date(year, month - 1, day, hours, minutes, parseInt(seconds)).getTime();
      }
      return date.getTime();
    } catch (e) {
      console.warn('Failed to parse timestamp:', timestamp);
      return 0; // Return earliest possible time if parsing fails
    }
  };

  const filteredLogs = useMemo(() => {
    return logs.filter(entry => {
      // Skip if user filters are active and this entry doesn't match any selected user
      if (userFilters.length > 0) {
        const entryUser = entry.UserId || entry.UserKey;
        if (!entryUser || !userFilters.includes(entryUser)) return false;
      }

      // Skip if workload filters are active and this entry doesn't match any selected workload
      if (workloadFilters.length > 0) {
        if (!entry.Workload || !workloadFilters.includes(entry.Workload)) return false;
      }

      // Skip if operation filters are active and this entry doesn't match any selected operation
      if (operationFilters.length > 0) {
        if (!entry.Operation || !operationFilters.includes(entry.Operation)) return false;
      }

      // Skip if IP filters are active and this entry doesn't match any selected IP
      if (ipFilters.length > 0) {
        const entryIP = entry.ClientIP || entry.ClientIPAddress;
        if (!entryIP || !ipFilters.includes(entryIP)) return false;
      }

      // Skip if correlation filter is active and this entry doesn't match
      if (correlationFilter) {
        if (!entry.CorrelationId || !entry.CorrelationId.includes(correlationFilter)) return false;
      }

      // Skip if showOnlyRisky is true and this entry isn't risky
      if (showOnlyRisky) {
        if (!riskyOps.includes(entry.Operation)) return false;
      }

      // Skip if search term is active and this entry doesn't match
      if (searchTerm) {
        const searchLower = searchTerm.toLowerCase();
        return (
          (entry.UserId?.toLowerCase().includes(searchLower) || entry.UserKey?.toLowerCase().includes(searchLower)) ||
          entry.Operation?.toLowerCase().includes(searchLower) ||
          entry.Workload?.toLowerCase().includes(searchLower) ||
          entry.Subject?.toLowerCase().includes(searchLower) ||
          entry.MessageId?.toLowerCase().includes(searchLower) ||
          entry.CorrelationId?.toLowerCase().includes(searchLower) ||
          (entry.ClientIP || entry.ClientIPAddress)?.toLowerCase().includes(searchLower) ||
          entry.FileName?.toLowerCase().includes(searchLower) ||
          entry.UserAgent?.toLowerCase().includes(searchLower)
        );
      }

      // Skip if entry is from NT AUTHORITY\SYSTEM
      const entryUser = entry.UserId || entry.UserKey;
      if (entryUser && entryUser.includes('NT AUTHORITY\\SYSTEM')) return false;

      return true;
    }).sort((a, b) => {
      const aTime = parseTimestamp(a.CreationDate || a.TimeGenerated || '');
      const bTime = parseTimestamp(b.CreationDate || b.TimeGenerated || '');
      return bTime - aTime; // Sort in descending order (newest first)
    });
  }, [logs, userFilters, workloadFilters, operationFilters, ipFilters, correlationFilter, showOnlyRisky, searchTerm]);

  const visibleLogs = filteredLogs.slice(0, visibleCount)

  const resetFilters = () => {
    setUserFilters([])
    setWorkloadFilters([])
    setOperationFilters([])
    setIpFilters([])
    setCorrelationFilter("")
    setShowOnlyRisky(false)
    setSearchTerm("")
    setVisibleCount(100)
  }

  // Toggle a filter in a multi-select array
  const toggleFilter = (type: "user" | "workload" | "operation" | "ip", value: string) => {
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
      case "ip":
        setIpFilters((prev) => (prev.includes(value) ? prev.filter((item) => item !== value) : [...prev, value]))
        break
    }
  }

  // Check if any filters are active
  const hasActiveFilters =
    userFilters.length > 0 ||
    workloadFilters.length > 0 ||
    operationFilters.length > 0 ||
    ipFilters.length > 0 ||
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

  const downloadInternetMessageIds = () => {
    // Create an interface for message statistics
    interface MessageStats {
      readBy: Set<string>;
      readAt: string[];
      subject: string;
      workload: string;
      clientIP: string;
      folderPath: string;
      sizeInBytes: number;
    }

    // Use a Record instead of Map
    const messageStats: Record<string, MessageStats> = {};

    // Process each log entry
    logs.forEach(entry => {
      try {
        // Parse AuditData if it's a string
        let auditData: any = entry.AuditData;
        if (typeof auditData === 'string') {
          auditData = JSON.parse(auditData);
        }

        // Look for InternetMessageId in the nested structure
        if (auditData.Folders && Array.isArray(auditData.Folders)) {
          auditData.Folders.forEach((folder: any) => {
            if (folder.FolderItems && Array.isArray(folder.FolderItems)) {
              folder.FolderItems.forEach((item: any) => {
                if (item.InternetMessageId && item.InternetMessageId.match(/<[^>]+@[^>]+>/)) {
                  const messageId = item.InternetMessageId;
                  if (!messageStats[messageId]) {
                    messageStats[messageId] = {
                      readBy: new Set<string>(),
                      readAt: [],
                      subject: entry.Subject || 'N/A',
                      workload: entry.Workload || 'Unknown',
                      clientIP: entry.ClientIP || 'N/A',
                      folderPath: folder.Path || 'Unknown',
                      sizeInBytes: item.SizeInBytes || 0
                    };
                  }

                  // Add user who read the message
                  if (entry.UserId || entry.UserKey) {
                    messageStats[messageId].readBy.add(entry.UserId || entry.UserKey);
                  }

                  // Add read timestamp
                  if (entry.TimeGenerated) {
                    messageStats[messageId].readAt.push(entry.TimeGenerated);
                  }
                }
              });
            }
          });
        }
      } catch (e) {
        console.warn("Failed to process log entry:", e);
      }
    });

    // Convert to CSV format
    const csvRows = ['InternetMessageId,Subject,Workload,Read By,Read Timestamps,Client IP,Folder Path,Size (Bytes)'];
    Object.entries(messageStats).forEach(([messageId, data]) => {
      const readBy = Array.from(data.readBy).join('; ');
      const readAt = data.readAt.join('; ');
      csvRows.push(`"${messageId}","${data.subject}","${data.workload}","${readBy}","${readAt}","${data.clientIP}","${data.folderPath}",${data.sizeInBytes}`);
    });

    // Create and download the file
    const csvContent = csvRows.join('\n');
    const blob = new Blob([csvContent], { type: 'text/csv;charset=utf-8;' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = 'mail-activity.csv';
    a.click();
    URL.revokeObjectURL(url);
  };

  const exportIPsToMap = async () => {
    // Define login operations
    const loginOperations = [
      "UserLoggedIn",
      "SignIn",
      "UserLoginFailed",
      "UserLoginSuccess"
    ];

    // Create a Set to store unique IPs
    const uniqueIPs = new Set<string>();

    // Process each log entry
    logs.forEach(entry => {
      if (!loginOperations.includes(entry.Operation)) return;
      
      if (userFilters.length > 0) {
        const entryUser = entry.UserId || entry.UserKey;
        if (!entryUser || !userFilters.includes(entryUser)) return;
      }

      const ip = entry.ClientIP;
      if (ip) {
        uniqueIPs.add(ip);
      }
    });

    // Convert IPs to newline-separated string
    const ipList = Array.from(uniqueIPs).join('\n');

    try {
      // Try the API endpoint first
      const response = await fetch('https://ipinfo.io/tools/map/cli', {
        method: 'POST',
        headers: {
          'Content-Type': 'text/plain',
        },
        body: ipList
      });

      const data = await response.json();
      
      if (data.reportUrl) {
        window.open(data.reportUrl, '_blank');
        return;
      }
      throw new Error('No report URL in response');
    } catch (error) {
      console.error('API call failed, falling back to form submission:', error);
      
      // Fallback to form submission
      const form = document.createElement('form');
      form.method = 'POST';
      form.action = 'https://ipinfo.io/tools/map';
      form.target = '_blank';

      const input = document.createElement('input');
      input.type = 'hidden';
      input.name = 'ips';
      input.value = ipList;

      form.appendChild(input);
      document.body.appendChild(form);
      form.submit();
      document.body.removeChild(form);
    }
  };

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

  const sortTimelineEvents = (events: TimelineEntry[], ascending: boolean) => {
    return [...events].sort((a, b) => {
      const aTime = parseTimestamp(a.timestamp);
      const bTime = parseTimestamp(b.timestamp);
      return ascending ? aTime - bTime : bTime - aTime;
    });
  };

  // Add function to add event to timeline
  const addToTimeline = (entry: LogEntry, title: string, description: string, type: 'info' | 'warning' | 'error' = 'info') => {
    const newEvent: TimelineEntry = {
      id: Math.random().toString(36).substr(2, 9),
      // Ensure consistent timestamp format
      timestamp: entry.CreationDate || entry.TimeGenerated || new Date().toISOString(),
      title,
      description,
      type,
      logEntry: entry,
      note: ''
    };

    setTimelineEvents(prevEvents => {
      // Check if an event with the same timestamp, title, and description already exists
      const isDuplicate = prevEvents.some(event => 
        event.timestamp === newEvent.timestamp && 
        event.title === newEvent.title && 
        event.description === newEvent.description
      );

      if (isDuplicate) {
        showNotification('This event is already in the timeline');
        return prevEvents;
      }

      const updatedEvents = [...prevEvents, newEvent];
      return sortTimelineEvents(updatedEvents, timelineSortAsc);
    });
  };

  // Add function to remove event from timeline
  const removeFromTimeline = (id: string) => {
    setTimelineEvents(prev => prev.filter(event => event.id !== id));
  };

  // Add function to show toast
  const showNotification = (message: string) => {
    setToastMessage(message);
    setShowToast(true);
    setTimeout(() => setShowToast(false), 3000); // Hide after 3 seconds
  };

  // Add function to update note
  const updateTimelineNote = (id: string, note: string) => {
    setTimelineEvents(prev => prev.map(event => 
      event.id === id ? { ...event, note } : event
    ));
  };

  const exportInvestigationTimeline = () => {
    // Create an HTML report
    let report = `
      <!DOCTYPE html>
      <html>
        <head>
          <title>Investigation Timeline Report</title>
          <style>
            body {
              font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, "Helvetica Neue", Arial, sans-serif;
              line-height: 1.6;
              color: #1a1a1a;
              max-width: 800px;
              margin: 0 auto;
              padding: 2rem;
            }
            .case-info {
              margin-bottom: 2rem;
              padding: 1rem;
              background-color: #f8fafc;
              border: 1px solid #e2e8f0;
              border-radius: 0.5rem;
            }
            .case-info h2 {
              color: #1e293b;
              margin-bottom: 1rem;
              padding-bottom: 0.5rem;
              border-bottom: 2px solid #e2e8f0;
            }
            .case-info table {
              width: 100%;
              border-collapse: collapse;
            }
            .case-info th, .case-info td {
              padding: 0.75rem;
              text-align: left;
              border-bottom: 1px solid #e2e8f0;
            }
            .case-info th {
              width: 30%;
              background-color: #f8fafc;
              font-weight: 600;
            }
            .timeline-event {
              margin-bottom: 2rem;
              padding: 1rem;
              border-left: 3px solid #3b82f6;
              background-color: #f8fafc;
              border-radius: 0.375rem;
            }
            .timeline-event.error {
              border-left-color: #ef4444;
              background-color: #fef2f2;
            }
            .timeline-event.warning {
              border-left-color: #f59e0b;
              background-color: #fffbeb;
            }
            .timeline-event h3 {
              margin: 0 0 0.5rem 0;
              color: #1e293b;
            }
            .timeline-event p {
              margin: 0;
              color: #475569;
            }
            .timestamp {
              font-size: 0.875rem;
              color: #64748b;
              margin-top: 0.5rem;
            }
            .analyst-note {
              margin-top: 1rem;
              padding: 1rem;
              background-color: #eff6ff;
              border: 1px solid #bfdbfe;
              border-radius: 0.375rem;
            }
            .analyst-note h4 {
              margin: 0 0 0.5rem 0;
              color: #1e40af;
              font-size: 0.875rem;
            }
            .analyst-note p {
              margin: 0;
              color: #1e40af;
              white-space: pre-wrap;
            }
            .header {
              text-align: center;
              margin-bottom: 2rem;
            }
            .header h1 {
              color: #1e293b;
              margin-bottom: 0.5rem;
            }
            .header p {
              color: #64748b;
            }
            .footer {
              text-align: center;
              margin-top: 2rem;
              padding-top: 1rem;
              border-top: 1px solid #e2e8f0;
              color: #64748b;
            }
            .report-section {
              margin-bottom: 2rem;
              padding: 1rem;
              background-color: #ffffff;
              border: 1px solid #e2e8f0;
              border-radius: 0.5rem;
            }
            .report-section h2 {
              color: #1e293b;
              margin-bottom: 1rem;
              padding-bottom: 0.5rem;
              border-bottom: 2px solid #e2e8f0;
            }
            table {
              width: 100%;
              border-collapse: collapse;
              margin: 1rem 0;
            }
            th, td {
              padding: 0.75rem;
              text-align: left;
              border-bottom: 1px solid #e2e8f0;
            }
            th {
              background-color: #f8fafc;
              font-weight: 600;
            }
            tr:hover {
              background-color: #f8fafc;
            }
            .inbox-rule {
              margin-top: 1rem;
              padding: 1rem;
              background-color: #fffbeb;
              border: 1px solid #fef3c7;
              border-radius: 0.375rem;
            }
            .inbox-rule h4 {
              color: #92400e;
              margin: 0 0 0.5rem 0;
            }
            .inbox-rule ul {
              margin: 0;
              padding-left: 1.5rem;
            }
            .inbox-rule li {
              margin: 0.25rem 0;
            }
          </style>
        </head>
        <body>
          <div class="header">
            <h1>Investigation Timeline Report</h1>
            <p>Generated on ${new Date().toLocaleString()}</p>
          </div>

          <div class="case-info">
            <h2>Case Information</h2>
            <table>
              <tr>
                <th>Case Name:</th>
                <td>${caseInfo.caseName || 'Not specified'}</td>
              </tr>
              <tr>
                <th>Case Number:</th>
                <td>${caseInfo.caseNumber || 'Not specified'}</td>
              </tr>
              <tr>
                <th>Analyst:</th>
                <td>${caseInfo.analystName || 'Not specified'}</td>
              </tr>
              <tr>
                <th>Analysis Date:</th>
                <td>${caseInfo.analysisDate || 'Not specified'}</td>
              </tr>
              <tr>
                <th>Company:</th>
                <td>${caseInfo.company || 'Not specified'}</td>
              </tr>
            </table>
          </div>
    `;

    // Add report sections if any exist
    if (reportSections.length > 0) {
      report += `<div class="report-section">
        <h2>Analysis Sections</h2>
        ${reportSections.map(section => {
          // Convert markdown table format to HTML table
          const content = section.content.includes('| Field | Value |')
            ? section.content
                .split('\n')
                .filter(line => line.trim() && !line.includes('---'))
                .map(line => {
                  if (line.startsWith('###')) {
                    return `<h3>${line.replace('###', '').trim()}</h3>`;
                  }
                  const [field, value] = line.split('|').filter(s => s.trim());
                  if (field && value) {
                    return `<tr><th>${field.trim()}</th><td>${value.trim()}</td></tr>`;
                  }
                  return '';
                })
                .join('\n')
              : section.content;

            return `
              <div class="timeline-event">
                <h3>${section.title}</h3>
                <div>
                  ${content.includes('<tr>') 
                    ? `<table class="w-full border-collapse">
                        ${content}
                      </table>`
                    : content}
                </div>
              </div>
            `;
          }).join('')}
      </div>`;
    }

    // Add timeline events
    report += `<div class="report-section">
      <h2>Timeline Events</h2>
      ${timelineEvents.map(event => {
        let eventContent = `
          <div class="timeline-event ${event.type}">
            <h3>${event.title}</h3>
            <p>${event.description}</p>
        `;

        // Add inbox rule details if present
        if (event.logEntry.RuleDetails) {
          eventContent += `
            <div class="inbox-rule">
              <h4>Inbox Rule Details</h4>
              <p><strong>Rule Name:</strong> ${event.logEntry.RuleDetails.Name}</p>
              <p><strong>Status:</strong> ${event.logEntry.RuleDetails.Enabled ? 'Enabled' : 'Disabled'}</p>
              ${event.logEntry.RuleDetails.Priority ? `<p><strong>Priority:</strong> ${event.logEntry.RuleDetails.Priority}</p>` : ''}
              ${event.logEntry.RuleDetails.StopProcessingRules ? '<p><strong>Stop Processing Rules:</strong> Yes</p>' : ''}
              ${event.logEntry.RuleDetails.Conditions.length > 0 ? `
                <p><strong>Conditions:</strong></p>
                <ul>
                  ${event.logEntry.RuleDetails.Conditions.map(condition => `<li>${condition}</li>`).join('')}
                </ul>
              ` : ''}
              ${event.logEntry.RuleDetails.Actions.length > 0 ? `
                <p><strong>Actions:</strong></p>
                <ul>
                  ${event.logEntry.RuleDetails.Actions.map(action => `<li>${action}</li>`).join('')}
                </ul>
              ` : ''}
            </div>
          `;
        }

        eventContent += `
            <div class="timestamp">${formatDate(event.timestamp)}</div>
            ${event.note ? `
              <div class="analyst-note">
                <h4>Analyst Note</h4>
                <p>${event.note}</p>
              </div>
            ` : ''}
          </div>
        `;

        return eventContent;
      }).join('')}
    </div>`;

    report += `
      <div class="footer">
        <p>Generated by UAL Timeline Builder</p>
      </div>
    </body>
    </html>
    `;

    // Create and download the file
    const blob = new Blob([report], { type: 'text/html' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = `investigation-timeline-${new Date().toISOString().split('T')[0]}.html`;
    a.click();
    URL.revokeObjectURL(url);
  };

  // Add after exportInvestigationTimeline function
  const saveTimeline = () => {
    const timelineData = {
      version: '1.0',
      exportDate: new Date().toISOString(),
      events: sortTimelineEvents(timelineEvents, timelineSortAsc)
    };
    
    const blob = new Blob([JSON.stringify(timelineData, null, 2)], { type: 'application/json' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = `investigation-timeline-${new Date().toISOString().split('T')[0]}.json`;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);
  };

  const loadTimeline = (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;

    const reader = new FileReader();
    reader.onload = (event) => {
      try {
        const content = event.target?.result as string;
        const timelineData = JSON.parse(content);
        
        if (timelineData.version === '1.0' && Array.isArray(timelineData.events)) {
          const sortedEvents = sortTimelineEvents(timelineData.events, timelineSortAsc);
          setTimelineEvents(sortedEvents);
          showNotification('Timeline loaded successfully');
        } else {
          showNotification('Invalid timeline file format');
        }
      } catch (error) {
        console.error('Error loading timeline:', error);
        showNotification('Error loading timeline file');
      }
    };
    reader.readAsText(file);
  };

  // Add a button to toggle sort direction in the timeline header
  const toggleSortDirection = () => {
    setTimelineSortAsc(prev => !prev);
    setTimelineEvents(prev => sortTimelineEvents(prev, !timelineSortAsc));
  };

  const analyzeAuthenticationBaseline = () => {
    // Create an object to store authentication statistics per IP and user
    const authStats: Record<string, AuthenticationStats> = {};
    
    // List of successful authentication operations
    const successfulAuthOperations = [
      'UserLoggedIn',
      'UserLoginSuccess',
      'TeamsSignIn',
      'ConsoleSignin',
      'AzurePortalSignin'
    ];

    // Process each log entry
    logs.forEach(entry => {
      if (!successfulAuthOperations.includes(entry.Operation)) {
        return;
      }

      const userId = entry.UserId;
      const clientIP = entry.ClientIP || entry.ClientIPAddress || 
        (typeof entry.AuditData === 'string' ? 
          JSON.parse(entry.AuditData)?.ClientIPAddress : 
          entry.AuditData?.ClientIPAddress);

      if (!userId || !clientIP) {
        return;
      }

      const key = `${clientIP}|${userId}`;
      if (!authStats[key]) {
        authStats[key] = {
          ip: clientIP,
          userId: userId,
          count: 0,
          firstSeen: entry.CreationDate,
          lastSeen: entry.CreationDate,
          operations: new Set<string>()
        };
      }

      const stats = authStats[key];
      stats.count += 1;
      stats.operations.add(entry.Operation);
      
      // Update timestamps
      if (new Date(entry.CreationDate) < new Date(stats.firstSeen)) {
        stats.firstSeen = entry.CreationDate;
      }
      if (new Date(entry.CreationDate) > new Date(stats.lastSeen)) {
        stats.lastSeen = entry.CreationDate;
      }
    });

    // Convert object to array and sort by authentication count
    const sortedData = Object.values(authStats)
      .sort((a, b) => b.count - a.count);
    
    setAuthBaselineData(sortedData);
    setShowAuthBaseline(true);
  };

  // Add modal close handler
  const closeAuthBaselineModal = () => {
    setShowAuthBaseline(false);
  };

  const handleCaseInfoChange = useCallback((e: React.ChangeEvent<HTMLInputElement>) => {
    const { name, value } = e.target;
    setCaseInfo(prev => ({
      ...prev,
      [name]: value
    }));
  }, []);

  const handleCaseInfoSubmit = (e: React.FormEvent) => {
    e.preventDefault();
    setShowCaseInfoModal(false);
    showNotification("Case information updated");
  };

  const handleDragEnter = (e: React.DragEvent) => {
    e.preventDefault()
    e.stopPropagation()
    setIsDragging(true)
  }

  const handleDragLeave = (e: React.DragEvent) => {
    e.preventDefault()
    e.stopPropagation()
    setIsDragging(false)
  }

  const handleDragOver = (e: React.DragEvent) => {
    e.preventDefault()
    e.stopPropagation()
  }

  const handleDrop = (e: React.DragEvent) => {
    e.preventDefault()
    e.stopPropagation()
    setIsDragging(false)

    const files = Array.from(e.dataTransfer.files)
    if (files.length === 0) return

    // Check if all files are CSV
    if (!files.every(file => file.name.endsWith(".csv"))) {
      alert("Only .csv files are supported.")
      return
    }

    setLoading(true)
    setFileNames(files.map(f => f.name))

    // Process files sequentially
    let allLogs: LogEntry[] = []
    let processedCount = 0

    const processFile = (file: File) => {
      Papa.parse(file, {
        header: true,
        skipEmptyLines: true,
        complete: (results) => {
          const data = results.data.map((entry: LogEntry) => {
            let auditData: AuditData = {}
            try {
              auditData = typeof entry.AuditData === 'string' ? JSON.parse(entry.AuditData) : entry.AuditData
            } catch (err) {
              console.warn("Failed to parse AuditData", err)
            }
            let ruleDetails = null
            if (entry.Operation === "UpdateInboxRule" || entry.Operation === "New-InboxRule" || entry.Operation === "Remove-InboxRule") {
              try {
                if (auditData.ModifiedProperties) {
                  ruleDetails = extractRuleDetails(auditData.ModifiedProperties)
                } else if (auditData.Parameters) {
                  ruleDetails = extractRuleDetailsFromParameters(auditData.Parameters)
                }
              } catch (err) {
                console.warn("Failed to parse rule details", err)
              }
            }

            // Parse User Agent if available
            let userAgentInfo: UserAgentInfo | undefined;
            if (auditData.UserAgent) {
              userAgentInfo = parseUserAgent(auditData.UserAgent);
            }

            return {
              ...entry,
              FileName: auditData.ObjectId ?? "",
              Subject: auditData.Subject ?? "",
              MessageId: auditData.MessageId ?? auditData.InternetMessageId ?? "",
              TimeGenerated: entry.CreationDate,
              ClientIP: auditData.ClientIP ?? auditData.ClientIPAddress ?? "",
              CorrelationId: auditData.CorrelationId ?? auditData.CorrelationID ?? "",
              ModifiedProperties: auditData.ModifiedProperties
                ? JSON.stringify(auditData.ModifiedProperties, null, 2)
                : "N/A",
              Workload: auditData.Workload ?? entry.Workload ?? "Unknown",
              AuditDataRaw: entry.AuditData,
              RuleDetails: ruleDetails ?? undefined,
              UserAgent: auditData.UserAgent,
              UserAgentInfo: userAgentInfo
            }
          })
          allLogs = [...allLogs, ...data]
          processedCount++

          // If all files are processed, update the state
          if (processedCount === files.length) {
            setLogs(allLogs)
            setLoading(false)
          }
        },
        error: (error) => {
          console.error("Error parsing file:", error)
          processedCount++
          if (processedCount === files.length) {
            setLogs(allLogs)
            setLoading(false)
          }
        }
      })
    }

    // Process each file
    files.forEach(processFile)
  }

  return (
    <div className="min-h-screen bg-white dark:bg-gray-900 p-4 md:p-6">
      {/* Toast Notification */}
      {showToast && (
        <div className="fixed bottom-4 right-4 bg-blue-600 text-white px-4 py-2 rounded-lg shadow-lg flex items-center gap-2 z-50 animate-fade-in-up">
          <Check className="h-4 w-4" />
          <span>{toastMessage}</span>
        </div>
      )}

      {/* Case Info Modal */}
      {showCaseInfoModal && (
        <div className="fixed inset-0 bg-black/50 flex items-center justify-center z-50">
          <div className="bg-white dark:bg-gray-800 rounded-lg shadow-xl max-w-md w-full mx-4">
            <div className="p-6">
              <h2 className="text-xl font-semibold text-slate-900 dark:text-white mb-4">Case Information</h2>
              <form onSubmit={handleCaseInfoSubmit}>
                <div className="space-y-4">
                  <div>
                    <label className="block text-sm font-medium text-slate-700 dark:text-slate-300 mb-1">
                      Case Name
                    </label>
                    <input
                      type="text"
                      name="caseName"
                      value={caseInfo.caseName}
                      onChange={handleCaseInfoChange}
                      className="w-full px-3 py-2 border border-slate-300 dark:border-slate-600 rounded-md bg-white dark:bg-gray-700 text-slate-900 dark:text-white"
                    />
                  </div>
                  <div>
                    <label className="block text-sm font-medium text-slate-700 dark:text-slate-300 mb-1">
                      Case Number
                    </label>
                    <input
                      type="text"
                      name="caseNumber"
                      value={caseInfo.caseNumber}
                      onChange={handleCaseInfoChange}
                      className="w-full px-3 py-2 border border-slate-300 dark:border-slate-600 rounded-md bg-white dark:bg-gray-700 text-slate-900 dark:text-white"
                    />
                  </div>
                  <div>
                    <label className="block text-sm font-medium text-slate-700 dark:text-slate-300 mb-1">
                      Analyst Name
                    </label>
                    <input
                      type="text"
                      name="analystName"
                      value={caseInfo.analystName}
                      onChange={handleCaseInfoChange}
                      className="w-full px-3 py-2 border border-slate-300 dark:border-slate-600 rounded-md bg-white dark:bg-gray-700 text-slate-900 dark:text-white"
                    />
                  </div>
                  <div>
                    <label className="block text-sm font-medium text-slate-700 dark:text-slate-300 mb-1">
                      Analysis Date
                    </label>
                    <input
                      type="date"
                      name="analysisDate"
                      value={caseInfo.analysisDate}
                      onChange={handleCaseInfoChange}
                      className="w-full px-3 py-2 border border-slate-300 dark:border-slate-600 rounded-md bg-white dark:bg-gray-700 text-slate-900 dark:text-white"
                    />
                  </div>
                  <div>
                    <label className="block text-sm font-medium text-slate-700 dark:text-slate-300 mb-1">
                      Company
                    </label>
                    <input
                      type="text"
                      name="company"
                      value={caseInfo.company}
                      onChange={handleCaseInfoChange}
                      className="w-full px-3 py-2 border border-slate-300 dark:border-slate-600 rounded-md bg-white dark:bg-gray-700 text-slate-900 dark:text-white"
                    />
                  </div>
                </div>
                <div className="mt-6 flex justify-end gap-3">
                  <button
                    type="button"
                    onClick={() => setShowCaseInfoModal(false)}
                    className="px-4 py-2 text-sm font-medium text-slate-700 bg-white border border-slate-300 rounded-md hover:bg-slate-50 dark:bg-slate-800 dark:text-slate-200 dark:border-slate-600 dark:hover:bg-slate-700"
                  >
                    Cancel
                  </button>
                  <button
                    type="submit"
                    className="px-4 py-2 text-sm font-medium text-white bg-blue-600 rounded-md hover:bg-blue-700"
                  >
                    Save
                  </button>
                </div>
              </form>
            </div>
          </div>
        </div>
      )}

      <div className="max-w-6xl mx-auto bg-white dark:bg-gray-900 shadow-xl rounded-xl border border-slate-200 dark:border-gray-700 overflow-hidden">
        {/* Header */}
        <div className="bg-gradient-to-r from-blue-600 to-indigo-600 p-6 text-white">
          <div className="flex items-center justify-between">
          <h1 className="text-3xl font-bold flex items-center gap-2">
            <Shield className="h-8 w-8" />
            UAL-Timeline-Builder (UTB)
          </h1>
            <div className="flex items-center gap-2">
              <button
                onClick={() => setShowTimeline(!showTimeline)}
                className="flex items-center gap-2 px-4 py-2 bg-white/10 hover:bg-white/20 rounded-lg transition-colors group relative"
                title=""
              >
                <Calendar className="h-5 w-5" />
                <span>Timeline {timelineEvents.length > 0 ? `(${timelineEvents.length})` : ''}</span>
                <div className="absolute hidden group-hover:block top-full left-1/2 transform -translate-x-1/2 mt-2 px-3 py-2 bg-gray-900 text-white text-xs rounded-lg whitespace-nowrap z-50">
                  Create and manage an investigation timeline
                  <div className="absolute bottom-full left-1/2 transform -translate-x-1/2">
                    <div className="border-8 border-transparent border-b-gray-900"></div>
                  </div>
                </div>
              </button>
              {renderDarkModeButton()}
            </div>
          </div>
          <p className="mt-2 opacity-90">Analyze and visualize Microsoft 365 Unified Audit Logs (or export to ndjson)</p>
          <p className="mt-2 text-sm font-medium text-yellow-200">
  ⚠️ This tool processes data entirely in your browser. No data is stored or transmitted to any server.
</p>
        </div>

        {/* Timeline View */}
        {showTimeline && (
          <div className="bg-white dark:bg-gray-800 rounded-lg shadow-sm border border-slate-200 dark:border-gray-700">
            <div className="p-4 border-b border-slate-200 dark:border-gray-700">
              <div className="flex items-center justify-between">
                <h2 className="text-lg font-semibold text-slate-900 dark:text-white">Investigation Timeline</h2>
                <div className="flex items-center gap-2">
                  <button
                    onClick={toggleSortDirection}
                    className="flex items-center gap-1 px-2 py-1 text-sm text-slate-600 dark:text-slate-300 hover:bg-slate-100 dark:hover:bg-slate-700 rounded"
                    title={timelineSortAsc ? "Oldest First" : "Newest First"}
                  >
                    <RefreshCw className="h-4 w-4" />
                    {timelineSortAsc ? "Oldest First" : "Newest First"}
                  </button>
                  <button
                    onClick={() => setShowCaseInfoModal(true)}
                    className="text-sm px-3 py-1.5 bg-slate-100 dark:bg-slate-800 text-slate-700 dark:text-slate-300 rounded-md hover:bg-slate-200 dark:hover:bg-slate-700 flex items-center gap-1.5"
                  >
                    <FileText className="h-4 w-4" />
                    Case Info
                  </button>
                  <button
                    onClick={exportInvestigationTimeline}
                    className="text-sm px-3 py-1.5 bg-blue-50 dark:bg-blue-900/30 text-blue-700 dark:text-blue-400 rounded-md hover:bg-blue-100 dark:hover:bg-blue-900/50 flex items-center gap-1.5"
                  >
                    <FileText className="h-4 w-4" />
                    Export Report
                  </button>
                  <button
                    onClick={saveTimeline}
                    className="text-sm px-3 py-1.5 bg-slate-50 dark:bg-gray-700 text-slate-700 dark:text-slate-300 rounded-md hover:bg-slate-100 dark:hover:bg-gray-600 flex items-center gap-1.5"
                  >
                    <Save className="h-4 w-4" />
                    Save Timeline
                  </button>
                  <label className="text-sm px-3 py-1.5 bg-slate-50 dark:bg-gray-700 text-slate-700 dark:text-slate-300 rounded-md hover:bg-slate-100 dark:hover:bg-gray-600 flex items-center gap-1.5 cursor-pointer">
                    <Upload className="h-4 w-4" />
                    Import Timeline
                    <input
                      type="file"
                      accept=".json"
                      onChange={loadTimeline}
                      className="hidden"
                    />
                  </label>
                </div>
              </div>
            </div>
            <div className="max-h-[600px] overflow-y-auto scrollbar-thin scrollbar-thumb-slate-300 dark:scrollbar-thumb-slate-600 scrollbar-track-slate-100 dark:scrollbar-track-slate-800 hover:scrollbar-thumb-slate-400 dark:hover:scrollbar-thumb-slate-500">
              <div className="p-4 space-y-4">
                {timelineEvents.length === 0 ? (
                  <div className="text-center py-8 text-slate-500 dark:text-slate-400">
                    <p className="mb-2">No events in timeline</p>
                    <p className="text-sm">Add events from the logs to build your investigation timeline</p>
                  </div>
                ) : (
                  timelineEvents.map((event) => (
                    <div
                      key={event.id}
                      className={`relative pl-8 pb-4 border-l-2 ${
                        event.type === 'error' ? 'border-red-500' :
                        event.type === 'warning' ? 'border-yellow-500' :
                        'border-blue-500'
                      }`}
                    >
                      <div className={`absolute left-[-5px] top-0 w-3 h-3 rounded-full ${
                        event.type === 'error' ? 'bg-red-500' :
                        event.type === 'warning' ? 'bg-yellow-500' :
                        'bg-blue-500'
                      }`} />
                      <div className="flex items-start justify-between">
                        <div>
                          <div className="text-sm font-medium">{event.title}</div>
                          <div className="text-sm text-slate-600 dark:text-slate-400">{event.description}</div>
                          <div className="text-xs text-slate-500 dark:text-slate-500 mt-1">
                            {formatDate(event.timestamp)}
                          </div>
                        </div>
                        <div className="flex items-start gap-2">
                          <button
                            onClick={() => setActiveNoteId(activeNoteId === event.id ? null : event.id)}
                            className={`text-sm px-2 py-1 rounded flex items-center gap-1 ${
                              event.note 
                                ? 'bg-blue-100 dark:bg-blue-900/30 text-blue-700 dark:text-blue-400 hover:bg-blue-200 dark:hover:bg-blue-900/50' 
                                : 'bg-slate-100 dark:bg-gray-800 text-slate-600 dark:text-slate-400 hover:bg-slate-200 dark:hover:bg-gray-700'
                            }`}
                            title={event.note || "Add a note"}
                          >
                            <PencilIcon className="h-3.5 w-3.5" />
                            {event.note ? 'Edit Note' : 'Add Note'}
                          </button>
                          <button
                            onClick={() => removeFromTimeline(event.id)}
                            className="text-slate-400 hover:text-slate-600 dark:hover:text-slate-300"
                          >
                            <X className="h-4 w-4" />
                          </button>
                        </div>
                      </div>
                      {activeNoteId === event.id && (
                        <div className="mt-2 relative">
                          <textarea
                            className="w-full px-2 py-1 text-sm bg-white dark:bg-gray-800 border border-blue-200 dark:border-blue-800 rounded-md focus:outline-none focus:ring-1 focus:ring-blue-500 dark:focus:ring-blue-400 resize-none shadow-sm"
                            placeholder="Add a note..."
                            rows={3}
                            value={event.note || ''}
                            onChange={(e) => updateTimelineNote(event.id, e.target.value)}
                            autoFocus
                          />
                          <button
                            onClick={() => setActiveNoteId(null)}
                            className="absolute top-1 right-1 text-slate-400 hover:text-slate-600 dark:hover:text-slate-300 p-1 rounded-full hover:bg-slate-100 dark:hover:bg-gray-700"
                          >
                            <X className="h-3 w-3" />
                          </button>
                        </div>
                      )}
                    </div>
                  ))
                )}
              </div>
            </div>
          </div>
        )}

        {/* Upload Section */}
        <div className="p-6 border-b border-slate-200 dark:border-gray-700">
          <div 
            className={`bg-slate-50 dark:bg-gray-800 rounded-xl p-6 border border-slate-200 dark:border-gray-700 transition-all hover:border-blue-300 dark:hover:border-blue-700 ${
              isDragging ? 'border-blue-500 dark:border-blue-500 bg-blue-50 dark:bg-blue-900/20' : ''
            }`}
            onDragEnter={handleDragEnter}
            onDragLeave={handleDragLeave}
            onDragOver={handleDragOver}
            onDrop={handleDrop}
          >
            <div className="flex flex-col md:flex-row md:items-center gap-4">
              <div className="flex-1">
                <h2 className="text-lg font-semibold mb-1">Upload Audit Log</h2>
                <p className="text-slate-500 dark:text-gray-400 text-sm">
                  {isDragging ? 'Drop your CSV files here' : 'Drag and drop or select CSV files containing UAL data'}
                </p>
              </div>
              <div className="flex-shrink-0">
                <label className="inline-flex items-center gap-2 px-4 py-2 bg-blue-600 hover:bg-blue-700 text-white rounded-lg cursor-pointer transition-colors">
                  <input type="file" accept=".csv" multiple onChange={handleFileUpload} className="hidden" />
                  <span>Select Files</span>
                </label>
              </div>
            </div>
            {fileNames.length > 0 && (
              <div className="mt-4 flex items-center gap-2 text-sm bg-blue-50 dark:bg-blue-900/30 text-blue-700 dark:text-blue-300 p-2 px-3 rounded-lg">
                <Info className="h-4 w-4" />
                <span>
                  Uploaded: <span className="font-medium">{fileNames.length} file{fileNames.length > 1 ? 's' : ''}</span>
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
                    <button
                      onClick={downloadInternetMessageIds}
                      className="flex items-center justify-center gap-1.5 px-3 py-1.5 bg-slate-800 hover:bg-slate-900 dark:bg-slate-700 dark:hover:bg-gray-600 text-white rounded-lg transition-colors text-sm"
                    >
                      <Mail className="h-3.5 w-3.5" />
                      <span>Export Mail Activity</span>
                      <span className="text-xs text-slate-300 dark:text-slate-400">{userFilters.length > 0 ? `${userFilters.length} users` : 'all users'}</span>
                    </button>
                    <button
                      onClick={downloadNDJSON}
                      className="flex items-center justify-center gap-1.5 px-3 py-1.5 bg-slate-800 hover:bg-slate-900 dark:bg-slate-700 dark:hover:bg-gray-600 text-white rounded-lg transition-colors text-sm"
                    >
                      <Download className="h-3.5 w-3.5" />
                      <span>Export to NDJSON</span>
                      <span className="text-xs text-slate-300 dark:text-slate-400">{userFilters.length > 0 ? `${userFilters.length} users` : 'all users'}</span>
                    </button>
                    
                    <button
                      onClick={() => analyzeAuthenticationBaseline()}
                      className="inline-flex items-center px-3 py-2 text-sm font-medium rounded-md text-blue-700 bg-blue-100 hover:bg-blue-200 dark:bg-blue-900/50 dark:text-blue-200 dark:hover:bg-blue-800/60 transition-colors duration-200 focus:outline-none focus:ring-2 focus:ring-offset-2 focus:ring-blue-500 dark:focus:ring-offset-slate-800"
                    >
                      <ChartBar className="h-4 w-4 mr-2 text-blue-600 dark:text-blue-400" />
                      Auth Baseline
                    </button>
                  </div>
                </div>

                {/* Authentication Baseline Analysis */}
                {showAuthBaseline && (
                  <div className="fixed inset-0 bg-black/50 backdrop-blur-sm flex items-center justify-center z-50">
                    <div className="bg-white dark:bg-slate-800 rounded-lg shadow-xl w-[95vw] max-w-6xl max-h-[90vh] flex flex-col">
                      <div className="flex items-center justify-between p-4 border-b border-slate-200 dark:border-slate-700">
                        <h2 className="text-xl font-semibold text-slate-900 dark:text-white">
                          Authentication Baseline Analysis
                        </h2>
                        <button
                          onClick={closeAuthBaselineModal}
                          className="text-slate-500 hover:text-slate-700 dark:text-slate-400 dark:hover:text-slate-200"
                        >
                          <X className="h-5 w-5" />
                        </button>
                      </div>
                      
                      <div className="flex-1 overflow-hidden p-4">
                        <div className="h-full overflow-auto">
                          <div className="min-w-full inline-block align-middle">
                            <table className="min-w-full divide-y divide-slate-200 dark:divide-slate-700">
                              <thead className="bg-slate-50 dark:bg-slate-800/50">
                                <tr>
                                  <th scope="col" className="w-[15%] px-4 py-3 text-left text-xs font-medium text-slate-500 dark:text-slate-400 uppercase tracking-wider">IP Address</th>
                                  <th scope="col" className="w-[20%] px-4 py-3 text-left text-xs font-medium text-slate-500 dark:text-slate-400 uppercase tracking-wider">User ID</th>
                                  <th scope="col" className="w-[10%] px-4 py-3 text-left text-xs font-medium text-slate-500 dark:text-slate-400 uppercase tracking-wider">Auth Count</th>
                                  <th scope="col" className="w-[15%] px-4 py-3 text-left text-xs font-medium text-slate-500 dark:text-slate-400 uppercase tracking-wider">First Seen</th>
                                  <th scope="col" className="w-[15%] px-4 py-3 text-left text-xs font-medium text-slate-500 dark:text-slate-400 uppercase tracking-wider">Last Seen</th>
                                  <th scope="col" className="w-[15%] px-4 py-3 text-left text-xs font-medium text-slate-500 dark:text-slate-400 uppercase tracking-wider">Operations</th>
                                  <th scope="col" className="w-[10%] px-4 py-3 text-left text-xs font-medium text-slate-500 dark:text-slate-400 uppercase tracking-wider">Actions</th>
                                </tr>
                              </thead>
                              <tbody className="divide-y divide-slate-200 dark:divide-slate-700">
                                {authBaselineData.map((data, index) => (
                                  <tr key={index} className="hover:bg-slate-50 dark:hover:bg-slate-700/50">
                                    <td className="px-4 py-3 text-sm text-slate-900 dark:text-white truncate max-w-[200px]" title={data.ip}>
                                      <button
                                        onClick={() => {
                                          toggleFilter("ip", data.ip);
                                          closeAuthBaselineModal();
                                        }}
                                        className="text-blue-600 dark:text-blue-400 hover:underline"
                                      >
                                        {data.ip}
                                      </button>
                                    </td>
                                    <td className="px-4 py-3 text-sm text-slate-900 dark:text-white truncate max-w-[250px]" title={data.userId}>{data.userId}</td>
                                    <td className="px-4 py-3 text-sm text-slate-900 dark:text-white">{data.count}</td>
                                    <td className="px-4 py-3 text-sm text-slate-900 dark:text-white">{formatDate(data.firstSeen)}</td>
                                    <td className="px-4 py-3 text-sm text-slate-900 dark:text-white">{formatDate(data.lastSeen)}</td>
                                    <td className="px-4 py-3 text-sm text-slate-900 dark:text-white">
                                      <div className="flex flex-wrap gap-1">
                                        {Array.from(data.operations).map((op, i) => (
                                          <span key={i} className="inline-flex items-center px-2 py-0.5 rounded text-xs font-medium bg-blue-100 text-blue-800 dark:bg-blue-900 dark:text-blue-200">
                                            {op}
                                          </span>
                                        ))}
                                      </div>
                                    </td>
                                    <td className="px-4 py-3 text-sm">
                                      <button
                                        onClick={() => {
                                          const baselineSection = {
                                            title: `Authentication Pattern - ${data.userId}`,
                                            content: `### \n\n| Field | Value |\n|-------|-------|\n| IP Address | ${data.ip} |\n| User ID | ${data.userId} |\n| Authentication Count | ${data.count} |\n| First Seen | ${data.firstSeen} |\n| Last Seen | ${data.lastSeen} |\n| Operations | ${Array.from(data.operations).join(', ')} |`
                                          };
                                          setReportSections([...reportSections, baselineSection]);
                                          closeAuthBaselineModal();
                                        }}
                                        className="inline-flex items-center px-3 py-2 text-sm font-medium rounded-md text-blue-700 bg-blue-100 hover:bg-blue-200 dark:bg-blue-900/50 dark:text-blue-200 dark:hover:bg-blue-800/60 transition-colors duration-200 focus:outline-none focus:ring-2 focus:ring-offset-2 focus:ring-blue-500 dark:focus:ring-offset-slate-800"
                                      >
                                        <FileText className="h-4 w-4 mr-2 text-blue-600 dark:text-blue-400" />
                                        Add to Report
                                      </button>
                                    </td>
                                  </tr>
                                ))}
                              </tbody>
                            </table>
                          </div>
                        </div>
                      </div>
                      
                      <div className="flex justify-end p-4 border-t border-slate-200 dark:border-slate-700">
                        <button
                          onClick={closeAuthBaselineModal}
                          className="px-4 py-2 text-sm font-medium text-slate-700 bg-white border border-slate-300 rounded-md hover:bg-slate-50 focus:outline-none focus:ring-2 focus:ring-offset-2 focus:ring-blue-500 dark:bg-slate-800 dark:text-slate-200 dark:border-slate-600 dark:hover:bg-slate-700"
                        >
                          Close
                        </button>
                      </div>
                    </div>
                  </div>
                )}

                {/* Filters */}
                <div className="bg-white/80 dark:bg-slate-900/40 backdrop-blur-sm border border-slate-200/60 dark:border-slate-800/60 rounded-xl p-6 mb-6">
                  <div className="flex items-center justify-between gap-4 mb-6">
                    <div className="flex items-center gap-2">
                      <h3 className="font-medium flex items-center gap-2 text-slate-200">
                        <Filter className="h-4 w-4" />
                        <span>Filters</span>
                      </h3>
                    </div>
                    
                    <div className="flex-1 max-w-md relative">
                      <Search className="absolute left-3 top-1/2 transform -translate-y-1/2 h-4 w-4 text-slate-400" />
                      <input
                        type="text"
                        placeholder="Search logs..."
                        onChange={(e) => debouncedSetSearchTerm(e.target.value)}
                        className="pl-9 pr-4 py-2 w-full rounded-lg bg-white dark:bg-slate-800/60 border border-slate-200 dark:border-slate-700/50 text-slate-900 dark:text-slate-200 placeholder:text-slate-500 dark:placeholder:text-slate-400 focus:ring-2 focus:ring-blue-500/40 focus:border-blue-500/40 outline-none transition-all"
                      />
                      {searchTerm && (
                        <button
                          onClick={() => setSearchTerm("")}
                          className="absolute right-3 top-1/2 transform -translate-y-1/2 text-slate-400 hover:text-slate-300 transition-colors"
                        >
                          <X className="h-4 w-4" />
                        </button>
                      )}
                    </div>

                    <button
                      onClick={resetFilters}
                      className={`text-sm flex items-center gap-1.5 px-3 py-1.5 rounded-lg transition-all ${
                        hasActiveFilters
                          ? "text-slate-200 hover:text-white bg-slate-700/50 hover:bg-slate-700/70"
                          : "text-slate-400 cursor-not-allowed"
                      }`}
                      disabled={!hasActiveFilters}
                    >
                      <RefreshCw className="h-3.5 w-3.5" />
                      Reset All
                    </button>
                  </div>

                  <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-4 gap-4">
                    {/* Multi-select User dropdown */}
                    <div ref={userDropdownRef} className="relative">
                      <label className="block text-sm font-medium text-slate-400 mb-1.5">Users</label>
                      <button
                        onClick={() => setUserDropdownOpen(!userDropdownOpen)}
                        className={`w-full flex items-center justify-between rounded-lg px-3 py-2 text-left transition-all ${
                          userFilters.length > 0
                            ? "bg-blue-500/10 border border-blue-500/20 text-blue-600 dark:text-blue-400"
                            : "bg-white dark:bg-slate-800/60 border border-slate-200 dark:border-slate-700/50 text-slate-900 dark:text-slate-200 hover:border-slate-300 dark:hover:border-slate-600/50"
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
                        <div className="absolute z-10 mt-1 w-full bg-slate-800 rounded-lg border border-slate-700 shadow-lg max-h-60 overflow-auto">
                          <div className="p-2 border-b border-slate-700 sticky top-0 bg-slate-800 z-10">
                            <div className="flex items-center justify-between mb-2">
                              <button
                                onClick={() => setUserFilters([])}
                                className="text-xs text-blue-400 hover:text-blue-300 transition-colors"
                              >
                                Clear all
                              </button>
                              <div className="relative flex-1 ml-4">
                                <Search className="absolute left-2 top-1/2 transform -translate-y-1/2 h-3 w-3 text-slate-400" />
                                <input
                                  type="text"
                                  placeholder="Search users..."
                                  value={userSearchTerm}
                                  onChange={(e) => setUserSearchTerm(e.target.value)}
                                  className="pl-7 pr-2 py-1 text-xs w-full rounded bg-slate-700/50 border border-slate-600/50 text-slate-200 placeholder-slate-400 focus:ring-1 focus:ring-blue-500/40 focus:border-blue-500/40 outline-none"
                                />
                              </div>
                            </div>
                          </div>
                          <div className="p-1">
                            {filteredUserOptions.map((user) => (
                              <div
                                key={user}
                                className="flex items-center px-3 py-2 hover:bg-slate-700/50 rounded cursor-pointer"
                                onClick={() => toggleFilter("user", user)}
                              >
                                <div className="mr-2 h-4 w-4 rounded border border-slate-600 flex items-center justify-center">
                                  {userFilters.includes(user) && (
                                    <Check className="h-3 w-3 text-blue-400" />
                                  )}
                                </div>
                                <span className="text-sm text-slate-200">{user}</span>
                              </div>
                            ))}
                          </div>
                        </div>
                      )}
                    </div>

                    {/* Multi-select Workload dropdown */}
                    <div ref={workloadDropdownRef} className="relative">
                      <label className="block text-sm font-medium text-slate-400 mb-1.5">Workloads</label>
                      <button
                        onClick={() => setWorkloadDropdownOpen(!workloadDropdownOpen)}
                        className={`w-full flex items-center justify-between rounded-lg px-3 py-2 text-left transition-all ${
                          workloadFilters.length > 0
                            ? "bg-blue-500/10 border border-blue-500/20 text-blue-600 dark:text-blue-400"
                            : "bg-white dark:bg-slate-800/60 border border-slate-200 dark:border-slate-700/50 text-slate-900 dark:text-slate-200 hover:border-slate-300 dark:hover:border-slate-600/50"
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
                        <div className="absolute z-10 mt-1 w-full bg-slate-800 rounded-lg border border-slate-700 shadow-lg max-h-60 overflow-auto">
                          <div className="p-2 border-b border-slate-700 sticky top-0 bg-slate-800 z-10">
                            <div className="flex items-center justify-between mb-2">
                              <button
                                onClick={() => setWorkloadFilters([])}
                                className="text-xs text-blue-400 hover:text-blue-300 transition-colors"
                              >
                                Clear all
                              </button>
                              <div className="relative flex-1 ml-4">
                                <Search className="absolute left-2 top-1/2 transform -translate-y-1/2 h-3 w-3 text-slate-400" />
                                <input
                                  type="text"
                                  placeholder="Search workloads..."
                                  value={workloadSearchTerm}
                                  onChange={(e) => setWorkloadSearchTerm(e.target.value)}
                                  className="pl-7 pr-2 py-1 text-xs w-full rounded bg-slate-700/50 border border-slate-600/50 text-slate-200 placeholder-slate-400 focus:ring-1 focus:ring-blue-500/40 focus:border-blue-500/40 outline-none"
                                />
                              </div>
                            </div>
                          </div>
                          <div className="p-1">
                            {filteredWorkloadOptions.map((workload) => (
                              <div
                                key={workload}
                                className="flex items-center px-3 py-2 hover:bg-slate-700/50 rounded cursor-pointer"
                                onClick={() => toggleFilter("workload", workload)}
                              >
                                <div className="mr-2 h-4 w-4 rounded border border-slate-600 flex items-center justify-center">
                                  {workloadFilters.includes(workload) && (
                                    <Check className="h-3 w-3 text-blue-400" />
                                  )}
                                </div>
                                <span className="text-sm text-slate-200">{workload}</span>
                              </div>
                            ))}
                          </div>
                        </div>
                      )}
                    </div>

                    {/* Multi-select Operation dropdown */}
                    <div ref={operationDropdownRef} className="relative">
                      <label className="block text-sm font-medium text-slate-400 mb-1.5">Operations</label>
                      <button
                        onClick={() => setOperationDropdownOpen(!operationDropdownOpen)}
                        className={`w-full flex items-center justify-between rounded-lg px-3 py-2 text-left transition-all ${
                          operationFilters.length > 0
                            ? "bg-blue-500/10 border border-blue-500/20 text-blue-600 dark:text-blue-400"
                            : "bg-white dark:bg-slate-800/60 border border-slate-200 dark:border-slate-700/50 text-slate-900 dark:text-slate-200 hover:border-slate-300 dark:hover:border-slate-600/50"
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
                        <div className="absolute z-10 mt-1 w-full bg-slate-800 rounded-lg border border-slate-700 shadow-lg max-h-60 overflow-auto">
                          <div className="p-2 border-b border-slate-700 sticky top-0 bg-slate-800 z-10">
                            <div className="flex items-center justify-between mb-2">
                              <button
                                onClick={() => setOperationFilters([])}
                                className="text-xs text-blue-400 hover:text-blue-300 transition-colors"
                              >
                                Clear all
                              </button>
                              <div className="relative flex-1 ml-4">
                                <Search className="absolute left-2 top-1/2 transform -translate-y-1/2 h-3 w-3 text-slate-400" />
                                <input
                                  type="text"
                                  placeholder="Search operations..."
                                  value={operationSearchTerm}
                                  onChange={(e) => setOperationSearchTerm(e.target.value)}
                                  className="pl-7 pr-2 py-1 text-xs w-full rounded bg-slate-700/50 border border-slate-600/50 text-slate-200 placeholder-slate-400 focus:ring-1 focus:ring-blue-500/40 focus:border-blue-500/40 outline-none"
                                />
                              </div>
                            </div>
                          </div>
                          <div className="p-1">
                            {filteredOperationOptions.map((operation) => (
                              <div
                                key={operation}
                                className="flex items-center px-3 py-2 hover:bg-slate-700/50 rounded cursor-pointer"
                                onClick={() => toggleFilter("operation", operation)}
                              >
                                <div className="mr-2 h-4 w-4 rounded border border-slate-600 flex items-center justify-center">
                                  {operationFilters.includes(operation) && (
                                    <Check className="h-3 w-3 text-blue-400" />
                                  )}
                                </div>
                                <span className="text-sm text-slate-200">
                                  {operation}
                                  {riskyOps.includes(operation) && (
                                    <span className="ml-2 inline-flex items-center text-xs font-medium text-red-400">
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

                    {/* Multi-select IP dropdown */}
                    <div ref={ipDropdownRef} className="relative">
                      <label className="block text-sm font-medium text-slate-400 mb-1.5">IP Addresses</label>
                      <button
                        onClick={() => setIpDropdownOpen(!ipDropdownOpen)}
                        className={`w-full flex items-center justify-between rounded-lg px-3 py-2 text-left transition-all ${
                          ipFilters.length > 0
                            ? "bg-blue-500/10 border border-blue-500/20 text-blue-600 dark:text-blue-400"
                            : "bg-white dark:bg-slate-800/60 border border-slate-200 dark:border-slate-700/50 text-slate-900 dark:text-slate-200 hover:border-slate-300 dark:hover:border-slate-600/50"
                        }`}
                      >
                        <span className="truncate">
                          {ipFilters.length === 0
                            ? "All IP Addresses"
                            : ipFilters.length === 1
                              ? ipFilters[0]
                              : `${ipFilters.length} IPs selected`}
                        </span>
                        <ChevronDown
                          className={`h-4 w-4 transition-transform ${ipDropdownOpen ? "rotate-180" : ""}`}
                        />
                      </button>

                      {ipDropdownOpen && (
                        <div className="absolute z-10 mt-1 w-full bg-slate-800 rounded-lg border border-slate-700 shadow-lg max-h-60 overflow-auto">
                          <div className="p-2 border-b border-slate-700 sticky top-0 bg-slate-800 z-10">
                            <div className="flex items-center justify-between mb-2">
                              <button
                                onClick={() => setIpFilters([])}
                                className="text-xs text-blue-400 hover:text-blue-300 transition-colors"
                              >
                                Clear all
                              </button>
                              <div className="relative flex-1 ml-4">
                                <Search className="absolute left-2 top-1/2 transform -translate-y-1/2 h-3 w-3 text-slate-400" />
                                <input
                                  type="text"
                                  placeholder="Search IPs..."
                                  value={ipSearchTerm}
                                  onChange={(e) => setIpSearchTerm(e.target.value)}
                                  className="pl-7 pr-2 py-1 text-xs w-full rounded bg-slate-700/50 border border-slate-600/50 text-slate-200 placeholder-slate-400 focus:ring-1 focus:ring-blue-500/40 focus:border-blue-500/40 outline-none"
                                />
                              </div>
                            </div>
                          </div>
                          <div className="p-1">
                            {filteredIpOptions.map((ip) => (
                              <div
                                key={ip}
                                className="flex items-center px-3 py-2 hover:bg-slate-700/50 rounded cursor-pointer"
                                onClick={() => toggleFilter("ip", ip)}
                              >
                                <div className="mr-2 h-4 w-4 rounded border border-slate-600 flex items-center justify-center">
                                  {ipFilters.includes(ip) && (
                                    <Check className="h-3 w-3 text-blue-400" />
                                  )}
                                </div>
                                <span className="text-sm text-slate-200">{ip}</span>
                              </div>
                            ))}
                          </div>
                        </div>
                      )}
                    </div>
                  </div>

                  {/* Risky Operations Toggle */}
                  <div className="mt-4">
                    <label
                      className={`flex items-center gap-2 w-full p-2 rounded-lg cursor-pointer transition-all ${
                        showOnlyRisky
                          ? "bg-blue-500/10 text-blue-600 dark:text-blue-400"
                          : "text-slate-700 dark:text-slate-200 hover:bg-slate-100 dark:hover:bg-slate-800/60"
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
                              ? "bg-blue-500 border-blue-500"
                              : "bg-white dark:bg-slate-800 border-slate-300 dark:border-slate-600"
                          }`}
                        ></div>
                        <svg
                          className={`absolute h-3 w-3 left-1 top-1 opacity-0 peer-checked:opacity-100 transition-opacity ${
                            showOnlyRisky ? "text-white" : "text-slate-900"
                          }`}
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

                  {/* Active Filter Chips */}
                  {hasActiveFilters && (
                    <div className="mt-4 flex flex-wrap gap-2">
                      {userFilters.map((user) => (
                        <div
                          key={`user-${user}`}
                          className="inline-flex items-center gap-2 px-3 py-1 bg-blue-500/10 text-blue-400 rounded-full text-sm border border-blue-500/20"
                        >
                          <span>User: {user}</span>
                          <button
                            onClick={() => toggleFilter("user", user)}
                            className="hover:bg-blue-400/10 rounded-full p-0.5 transition-colors"
                          >
                            <X className="h-3 w-3" />
                            <span className="sr-only">Remove user filter</span>
                          </button>
                        </div>
                      ))}

                      {workloadFilters.map((workload) => (
                        <div
                          key={`workload-${workload}`}
                          className="inline-flex items-center gap-2 px-3 py-1 bg-blue-500/10 text-blue-400 rounded-full text-sm border border-blue-500/20"
                        >
                          <span>Workload: {workload}</span>
                          <button
                            onClick={() => toggleFilter("workload", workload)}
                            className="hover:bg-blue-400/10 rounded-full p-0.5 transition-colors"
                          >
                            <X className="h-3 w-3" />
                            <span className="sr-only">Remove workload filter</span>
                          </button>
                        </div>
                      ))}

                      {operationFilters.map((operation) => (
                        <div
                          key={`operation-${operation}`}
                          className="inline-flex items-center gap-2 px-3 py-1 bg-blue-500/10 text-blue-400 rounded-full text-sm border border-blue-500/20"
                        >
                          <span>Operation: {operation}</span>
                          <button
                            onClick={() => toggleFilter("operation", operation)}
                            className="hover:bg-blue-400/10 rounded-full p-0.5 transition-colors"
                          >
                            <X className="h-3 w-3" />
                            <span className="sr-only">Remove operation filter</span>
                          </button>
                        </div>
                      ))}

                      {ipFilters.map((ip) => (
                        <div
                          key={`ip-${ip}`}
                          className="inline-flex items-center gap-2 px-3 py-1 bg-blue-500/10 text-blue-400 rounded-full text-sm border border-blue-500/20"
                        >
                          <span>IP: {ip}</span>
                          <button
                            onClick={() => toggleFilter("ip", ip)}
                            className="hover:bg-blue-400/10 rounded-full p-0.5 transition-colors"
                          >
                            <X className="h-3 w-3" />
                            <span className="sr-only">Remove IP filter</span>
                          </button>
                        </div>
                      ))}

                      {correlationFilter && (
                        <div className="inline-flex items-center gap-2 px-3 py-1 bg-blue-500/10 text-blue-400 rounded-full text-sm border border-blue-500/20">
                          <span>Correlation ID: {correlationFilter.substring(0, 10)}...</span>
                          <button
                            onClick={() => setCorrelationFilter("")}
                            className="hover:bg-blue-400/10 rounded-full p-0.5 transition-colors"
                          >
                            <X className="h-3 w-3" />
                            <span className="sr-only">Remove correlation filter</span>
                          </button>
                        </div>
                      )}

                      {showOnlyRisky && (
                        <div className="inline-flex items-center gap-2 px-3 py-1 bg-blue-500/10 text-blue-400 rounded-full text-sm border border-blue-500/20">
                          <span>Risky Operations Only</span>
                          <button
                            onClick={() => setShowOnlyRisky(false)}
                            className="hover:bg-blue-400/10 rounded-full p-0.5 transition-colors"
                          >
                            <X className="h-3 w-3" />
                            <span className="sr-only">Remove risky operations filter</span>
                          </button>
                        </div>
                      )}

                      {searchTerm && (
                        <div className="inline-flex items-center gap-2 px-3 py-1 bg-blue-500/10 text-blue-400 rounded-full text-sm border border-blue-500/20">
                          <span>Search: {searchTerm}</span>
                          <button
                            onClick={() => setSearchTerm("")}
                            className="hover:bg-blue-400/10 rounded-full p-0.5 transition-colors"
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

            <button
                                onClick={() => {
                                  const title = `${entry.Operation} by ${entry.UserId || entry.UserKey}`;
                                  let description = `Workload: ${entry.Workload}${entry.ClientIP ? `, IP: ${entry.ClientIP}` : ''}`;

                                  // Add operation-specific details
                                  if (entry.Operation === "Add user.") {
                                    try {
                                      const modifiedProps = typeof entry.ModifiedProperties === 'string' 
                                        ? JSON.parse(entry.ModifiedProperties)
                                        : entry.ModifiedProperties;
                                      
                                      const userPrincipalName = modifiedProps.find((prop: any) => prop.Name === "UserPrincipalName")?.NewValue;
                                      const displayName = modifiedProps.find((prop: any) => prop.Name === "DisplayName")?.NewValue;
                                      const accountEnabled = modifiedProps.find((prop: any) => prop.Name === "AccountEnabled")?.NewValue;
                                      
                                      description += `\nCreated user: ${displayName || userPrincipalName || 'Unknown'}`;
                                      description += `\nAccount Status: ${accountEnabled === "True" ? "Enabled" : "Disabled"}`;
                                    } catch (e) {
                                      console.warn("Failed to parse user creation details:", e);
                                    }
                                  } 
                                  else if (entry.Operation === "UpdateInboxRule" || entry.Operation === "New-InboxRule" || entry.Operation === "Remove-InboxRule") {
                                    if (entry.RuleDetails) {
                                      description += `\nRule Name: ${entry.RuleDetails.Name}`;
                                      if (entry.RuleDetails.Actions?.length) {
                                        description += `\nActions: ${entry.RuleDetails.Actions.join(", ")}`;
                                      }
                                      description += `\nStatus: ${entry.RuleDetails.Enabled ? "Enabled" : "Disabled"}`;
                                    }
                                  }
                                  else if (entry.Operation === "Add member to role.") {
                                    try {
                                      const modifiedProps = typeof entry.ModifiedProperties === 'string' 
                                        ? JSON.parse(entry.ModifiedProperties)
                                        : entry.ModifiedProperties;
                                      
                                      const roleDisplayName = modifiedProps.find((prop: any) => prop.Name === "Role.DisplayName")?.NewValue;
                                      const targetUser = entry.FileName;
                                      
                                      description += `\nRole: ${roleDisplayName || 'Unknown'}`;
                                      description += `\nAssigned to: ${targetUser || 'Unknown'}`;
                                    } catch (e) {
                                      console.warn("Failed to parse role assignment details:", e);
                                    }
                                  }
                                  else if (entry.Operation === "Update application – Certificates and secrets management " || 
                                          entry.Operation === "Add application." || 
                                          entry.Operation === "Update application.") {
                                    try {
                                      const modifiedProps = typeof entry.ModifiedProperties === 'string' 
                                        ? JSON.parse(entry.ModifiedProperties)
                                        : entry.ModifiedProperties;
                                      
                                      const keyDescriptionProp = modifiedProps.find((prop: any) => 
                                        prop.Name === "KeyDescription" || 
                                        prop.Name === "KeyDescriptions"
                                      );
                                      
                                      if (keyDescriptionProp) {
                                        const parseKeyList = (value: string) => {
                                          try {
                                            const parsed = JSON.parse(value);
                                            return parsed.map((item: string) => {
                                              const match = item.match(/KeyIdentifier=([^,]+),KeyType=([^,]+),KeyUsage=([^,]+),DisplayName=([^\]]+)/);
                                              if (match) {
                                                return {
                                                  keyId: match[1],
                                                  keyType: match[2],
                                                  keyUsage: match[3],
                                                  displayName: match[4]
                                                };
                                              }
                                              return null;
                                            }).filter(Boolean);
                                          } catch (e) {
                                            console.warn('Failed to parse key list:', e);
                                            return [];
                                          }
                                        };

                                        const oldKeys = parseKeyList(keyDescriptionProp.OldValue);
                                        const newKeys = parseKeyList(keyDescriptionProp.NewValue);
                                        const addedKeys = newKeys.filter((newKey: any) => 
                                          !oldKeys.some((oldKey: any) => oldKey.keyId === newKey.keyId)
                                        );
                                        const removedKeys = oldKeys.filter((oldKey: any) => 
                                          !newKeys.some((newKey: any) => newKey.keyId === oldKey.keyId)
                                        );

                                        if (addedKeys.length) {
                                          description += `\nAdded Keys: ${addedKeys.map((k: any) => k.displayName).join(", ")}`;
                                        }
                                        if (removedKeys.length) {
                                          description += `\nRemoved Keys: ${removedKeys.map((k: any) => k.displayName).join(", ")}`;
                                        }
                                      }
                                    } catch (e) {
                                      console.warn("Failed to parse certificate/secret changes:", e);
                                    }
                                  }

                                  addToTimeline(entry, title, description, isRisky ? 'warning' : 'info');
                                  showNotification(`Added ${entry.Operation} to timeline`);
                                }}
                                className="ml-auto flex items-center gap-1.5 px-3 py-1.5 text-sm bg-blue-600 hover:bg-blue-700 text-white rounded-lg transition-colors"
                              >
                                <Calendar className="h-3.5 w-3.5" />
                                <span>Add to Investigation</span>
                              </button>
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
  {/* Special display for inbox rules */}
  {(entry.Operation === "UpdateInboxRule" || entry.Operation === "New-InboxRule" || entry.Operation === "Remove-InboxRule") &&
                              entry.RuleDetails && (
                                <div className="mt-3 mb-3 bg-yellow-50 dark:bg-yellow-900/20 border border-yellow-200 dark:border-yellow-800 rounded-lg p-3">
                                  <h4 className="font-bold text-yellow-800 dark:text-yellow-300 mb-2 flex items-center gap-2">
                                    <AlertTriangle className="h-4 w-4" />
                                    Inbox Rule Details
                                  </h4>

                                  <div className="grid grid-cols-1 md:grid-cols-2 gap-3 text-sm">
                                    {entry.RuleDetails.Name && (
                                      <div>
                                        <span className="font-bold text-yellow-700 dark:text-yellow-400">
                                          Rule Name:
                                        </span>{" "}
                                        <span className="text-yellow-900 dark:text-yellow-200 font-normal">
                                          {entry.RuleDetails.Name}
                                        </span>
                                      </div>
                                    )}

                                    {entry.RuleDetails.Priority && (
                                      <div>
                                        <span className="font-bold text-yellow-700 dark:text-yellow-400">
                                          Priority:
                                        </span>{" "}
                                        <span className="text-yellow-900 dark:text-yellow-200 font-normal">
                                          {entry.RuleDetails.Priority}
                                        </span>
                                      </div>
                                    )}

                                    <div className="md:col-span-2">
                                      <span className="font-bold text-yellow-700 dark:text-yellow-400">Status:</span>{" "}
                                      <span
                                        className={`font-normal ${entry.RuleDetails.Enabled ? "text-green-600 dark:text-green-400" : "text-red-600 dark:text-red-400"}`}
                                      >
                                        {entry.RuleDetails.Enabled ? "Enabled" : "Disabled"}
                                      </span>
                                      {entry.RuleDetails.StopProcessingRules && (
                                        <span className="ml-2 text-red-600 dark:text-red-400 font-normal">
                                          (Stops processing other rules)
                                        </span>
                                      )}
                                    </div>

                                    {entry.RuleDetails.Conditions && entry.RuleDetails.Conditions.length > 0 && (
                                      <div className="md:col-span-2">
                                        <span className="font-bold text-yellow-700 dark:text-yellow-400 block mb-1">
                                          Conditions:
                                        </span>
                                        <ul className="list-disc pl-5 space-y-1 text-yellow-900 dark:text-yellow-200 font-normal">
                                          {entry.RuleDetails.Conditions.map((condition, i) => (
                                            <li key={i}>{condition}</li>
                                          ))}
                                        </ul>
                                      </div>
                                    )}

                                    {entry.RuleDetails.Actions && entry.RuleDetails.Actions.length > 0 && (
                                      <div className="md:col-span-2">
                                        <span className="font-bold text-yellow-700 dark:text-yellow-400 block mb-1">
                                          Actions:
                                        </span>
                                        <ul className="list-disc pl-5 space-y-1 text-yellow-900 dark:text-yellow-200 font-normal">
                                          {entry.RuleDetails.Actions.map((action, i) => (
                                            <li key={i}>{action}</li>
                                          ))}
                                        </ul>
                                      </div>
                                    )}
                                  </div>
                                </div>
                              )}

                              {/* Special display for role assignments */}
                              {entry.Operation === "Add member to role." && entry.ModifiedProperties && (
                                <div className="mt-3 mb-3 bg-yellow-50 dark:bg-yellow-900/20 border border-yellow-200 dark:border-yellow-800 rounded-lg p-3">
                                  <h4 className="font-bold text-yellow-800 dark:text-yellow-300 mb-2 flex items-center gap-2">
                                    <AlertTriangle className="h-4 w-4" />
                                    Role Assignment Details
                                  </h4>
                                  <div className="grid grid-cols-1 md:grid-cols-2 gap-3 text-sm">
                                    {(() => {
                                      let modifiedProps: any[] = [];
                                      try {
                                        if (typeof entry.ModifiedProperties === 'string') {
                                          modifiedProps = JSON.parse(entry.ModifiedProperties);
                                        } else {
                                          modifiedProps = entry.ModifiedProperties;
                                        }
                                      } catch (e) {
                                        console.warn('Failed to parse ModifiedProperties:', e);
                                        return null;
                                      }

                                      const roleDisplayName = modifiedProps.find(prop => prop.Name === "Role.DisplayName")?.NewValue;
                                      const roleObjectId = modifiedProps.find(prop => prop.Name === "Role.ObjectID")?.NewValue;
                                      const roleTemplateId = modifiedProps.find(prop => prop.Name === "Role.TemplateId")?.NewValue;
                                      const roleWellKnownName = modifiedProps.find(prop => prop.Name === "Role.WellKnownObjectName")?.NewValue;

                                      return (
                                        <>
                                          {roleDisplayName && (
                                            <div>
                                              <span className="font-bold text-yellow-700 dark:text-yellow-400">Role Name:</span>{" "}
                                              <span className="text-yellow-900 dark:text-yellow-200 font-normal">
                                                {roleDisplayName}
                                              </span>
                                            </div>
                                          )}
                                          {roleWellKnownName && (
                                            <div>
                                              <span className="font-bold text-yellow-700 dark:text-yellow-400">Well-Known Name:</span>{" "}
                                              <span className="text-yellow-900 dark:text-yellow-200 font-normal">
                                                {roleWellKnownName}
                                              </span>
                                            </div>
                                          )}
                                          {roleObjectId && (
                                            <div>
                                              <span className="font-bold text-yellow-700 dark:text-yellow-400">Role ID:</span>{" "}
                                              <span className="text-yellow-900 dark:text-yellow-200 font-normal break-all">
                                                {roleObjectId}
                                              </span>
                                            </div>
                                          )}
                                          {roleTemplateId && (
                                            <div>
                                              <span className="font-bold text-yellow-700 dark:text-yellow-400">Template ID:</span>{" "}
                                              <span className="text-yellow-900 dark:text-yellow-200 font-normal break-all">
                                                {roleTemplateId}
                                              </span>
                                            </div>
                                          )}
                                          <div className="md:col-span-2">
                                            <span className="font-bold text-yellow-700 dark:text-yellow-400">Target User:</span>{" "}
                                            <span className="text-yellow-900 dark:text-yellow-200 font-normal">
                                              {entry.FileName || "Unknown"}
                                            </span>
                                          </div>
                                        </>
                                      );
                                    })()}
                                  </div>
                                </div>
                              )}

                              {/* Special display for user creation */}
                              {entry.Operation === "Add user." && entry.ModifiedProperties && (
                                <div className="mt-3 mb-3 bg-yellow-50 dark:bg-yellow-900/20 border border-yellow-200 dark:border-yellow-800 rounded-lg p-3">
                                  <h4 className="font-bold text-yellow-800 dark:text-yellow-300 mb-2 flex items-center gap-2">
                                    <AlertTriangle className="h-4 w-4" />
                                    User Creation Details
                                  </h4>
                                  <div className="grid grid-cols-1 md:grid-cols-2 gap-3 text-sm">
                                    {(() => {
                                      let modifiedProps: any[] = [];
                                      try {
                                        if (typeof entry.ModifiedProperties === 'string') {
                                          modifiedProps = JSON.parse(entry.ModifiedProperties);
                                        } else {
                                          modifiedProps = entry.ModifiedProperties;
                                        }
                                      } catch (e) {
                                        console.warn('Failed to parse ModifiedProperties:', e);
                                        return null;
                                      }

                                      const userPrincipalName = modifiedProps.find(prop => prop.Name === "UserPrincipalName")?.NewValue;
                                      const displayName = modifiedProps.find(prop => prop.Name === "DisplayName")?.NewValue;
                                      const accountEnabled = modifiedProps.find(prop => prop.Name === "AccountEnabled")?.NewValue;
                                      const userType = modifiedProps.find(prop => prop.Name === "UserType")?.NewValue;
                                      const department = modifiedProps.find(prop => prop.Name === "Department")?.NewValue;
                                      const jobTitle = modifiedProps.find(prop => prop.Name === "JobTitle")?.NewValue;
                                      const office = modifiedProps.find(prop => prop.Name === "Office")?.NewValue;
                                      const usageLocation = modifiedProps.find(prop => prop.Name === "UsageLocation")?.NewValue;

                                      return (
                                        <>
                                          {userPrincipalName && (
                                            <div>
                                              <span className="font-bold text-yellow-700 dark:text-yellow-400">User Principal Name:</span>{" "}
                                              <span className="text-yellow-900 dark:text-yellow-200 font-normal">
                                                {userPrincipalName}
                                              </span>
                                            </div>
                                          )}
                                          {displayName && (
                                            <div>
                                              <span className="font-bold text-yellow-700 dark:text-yellow-400">Display Name:</span>{" "}
                                              <span className="text-yellow-900 dark:text-yellow-200 font-normal">
                                                {displayName}
                                              </span>
                                            </div>
                                          )}
                                          <div>
                                            <span className="font-bold text-yellow-700 dark:text-yellow-400">Account Status:</span>{" "}
                                            <span className={`font-normal ${accountEnabled === "True" ? "text-green-600 dark:text-green-400" : "text-red-600 dark:text-red-400"}`}>
                                              {accountEnabled === "True" ? "Enabled" : "Disabled"}
                                            </span>
                                          </div>
                                          {userType && (
                                            <div>
                                              <span className="font-bold text-yellow-700 dark:text-yellow-400">User Type:</span>{" "}
                                              <span className="text-yellow-900 dark:text-yellow-200 font-normal">
                                                {userType}
                                              </span>
                                            </div>
                                          )}
                                          {department && (
                                            <div>
                                              <span className="font-bold text-yellow-700 dark:text-yellow-400">Department:</span>{" "}
                                              <span className="text-yellow-900 dark:text-yellow-200 font-normal">
                                                {department}
                                              </span>
                                            </div>
                                          )}
                                          {jobTitle && (
                                            <div>
                                              <span className="font-bold text-yellow-700 dark:text-yellow-400">Job Title:</span>{" "}
                                              <span className="text-yellow-900 dark:text-yellow-200 font-normal">
                                                {jobTitle}
                                              </span>
                                            </div>
                                          )}
                                          {office && (
                                            <div>
                                              <span className="font-bold text-red-700 dark:text-red-400">Office:</span>{" "}
                                              <span className="text-red-900 dark:text-red-200 font-normal">
                                                {office}
                                              </span>
                                            </div>
                                          )}
                                          {usageLocation && (
                                            <div>
                                              <span className="font-bold text-red-700 dark:text-red-400">Usage Location:</span>{" "}
                                              <span className="text-red-900 dark:text-red-200 font-normal">
                                                {usageLocation}
                                              </span>
                                            </div>
                                          )}
                                        </>
                                      );
                                    })()}
                                  </div>
                                </div>
                              )}
                            <details className="group">
                              <summary className="cursor-pointer list-none flex items-center gap-1 text-sm font-medium text-slate-600 dark:text-gray-300 hover:text-blue-600 dark:hover:text-blue-400 transition-colors">
                                <ChevronDown className="h-4 w-4 transition-transform group-open:rotate-180" />
                                View Details
                              </summary>
                              <div className="mt-3 pt-3 border-t border-slate-200 dark:border-slate-800 text-sm space-y-3">
                                {/* App Installation Details */}
                                {entry.Operation === "AppInstalled" && (
                                  <div>
                                    <div className="mb-2 flex items-center gap-2 text-purple-700 dark:text-purple-300">
                                      <Monitor className="h-5 w-5" />
                                      <span className="font-semibold">App Installation Details</span>
                                    </div>
                                    <div className="bg-purple-50 dark:bg-purple-900/20 border border-purple-200 dark:border-purple-800 rounded-lg p-4">
                                      <div className="space-y-3">
                                        {(() => {
                                          let parsedAuditData: AuditData | null = null;
                                          try {
                                            parsedAuditData = typeof entry.AuditData === 'string' 
                                              ? JSON.parse(entry.AuditData)
                                              : entry.AuditData;
                                          } catch (e) {
                                            console.warn('Failed to parse AuditData:', e);
                                          }

                                          if (!parsedAuditData) return null;

                                          return (
                                            <>
                                              {parsedAuditData.AddOnName && (
                                                <div className="flex items-start gap-3 p-2 bg-white/50 dark:bg-black/10 rounded border border-purple-100 dark:border-purple-900">
                                                  <Monitor className="h-5 w-5 text-purple-600 dark:text-purple-400 flex-shrink-0 mt-0.5" />
                                                  <div>
                                                    <div className="text-sm font-medium text-purple-700 dark:text-purple-300">App Name</div>
                                                    <div className="text-sm text-slate-900 dark:text-slate-100">{parsedAuditData.AddOnName}</div>
                                                  </div>
                                                </div>
                                              )}
                                              {parsedAuditData.AppDistributionMode && (
                                                <div className="flex items-start gap-3 p-2 bg-white/50 dark:bg-black/10 rounded border border-purple-100 dark:border-purple-900">
                                                  <Globe className="h-5 w-5 text-purple-600 dark:text-purple-400 flex-shrink-0 mt-0.5" />
                                                  <div>
                                                    <div className="text-sm font-medium text-purple-700 dark:text-purple-300">Distribution Mode</div>
                                                    <div className="text-sm text-slate-900 dark:text-slate-100">{parsedAuditData.AppDistributionMode}</div>
                                                  </div>
                                                </div>
                                              )}
                                              {parsedAuditData.AzureADAppId && (
                                                <div className="flex items-start gap-3 p-2 bg-white/50 dark:bg-black/10 rounded border border-purple-100 dark:border-purple-900">
                                                  <Shield className="h-5 w-5 text-purple-600 dark:text-purple-400 flex-shrink-0 mt-0.5" />
                                                  <div>
                                                    <div className="text-sm font-medium text-purple-700 dark:text-purple-300">Azure AD App ID</div>
                                                    <div className="text-sm text-slate-900 dark:text-slate-100 break-all">{parsedAuditData.AzureADAppId}</div>
                                                  </div>
                                                </div>
                                              )}
                                              {parsedAuditData.ResourceSpecificApplicationPermissions && parsedAuditData.ResourceSpecificApplicationPermissions.length > 0 && (
                                                <div className="flex items-start gap-3 p-2 bg-white/50 dark:bg-black/10 rounded border border-purple-100 dark:border-purple-900">
                                                  <Shield className="h-5 w-5 text-purple-600 dark:text-purple-400 flex-shrink-0 mt-0.5" />
                                                  <div>
                                                    <div className="text-sm font-medium text-purple-700 dark:text-purple-300">App Permissions</div>
                                                    <div className="text-sm text-slate-900 dark:text-slate-100">
                                                      <ul className="list-disc pl-5 space-y-1">
                                                        {parsedAuditData.ResourceSpecificApplicationPermissions.map((permission, index) => (
                                                          <li key={index}>{permission}</li>
                                                        ))}
                                                      </ul>
                                                    </div>
                                                  </div>
                                                </div>
                                              )}
                                              {parsedAuditData.ChatThreadId && (
                                                <div className="flex items-start gap-3 p-2 bg-white/50 dark:bg-black/10 rounded border border-purple-100 dark:border-purple-900">
                                                  <Inbox className="h-5 w-5 text-purple-600 dark:text-purple-400 flex-shrink-0 mt-0.5" />
                                                  <div>
                                                    <div className="text-sm font-medium text-purple-700 dark:text-purple-300">Chat Thread ID</div>
                                                    <div className="text-sm text-slate-900 dark:text-slate-100 break-all">{parsedAuditData.ChatThreadId}</div>
                                                  </div>
                                                </div>
                                              )}
                                              {parsedAuditData.DeviceId && (
                                                <div className="flex items-start gap-3 p-2 bg-white/50 dark:bg-black/10 rounded border border-purple-100 dark:border-purple-900">
                                                  <Laptop className="h-5 w-5 text-purple-600 dark:text-purple-400 flex-shrink-0 mt-0.5" />
                                                  <div>
                                                    <div className="text-sm font-medium text-purple-700 dark:text-purple-300">Device ID</div>
                                                    <div className="text-sm text-slate-900 dark:text-slate-100 break-all">{parsedAuditData.DeviceId}</div>
                                                  </div>
                                                </div>
                                              )}
                                            </>
                                          );
                                        })()}
                                      </div>
                                    </div>
                                  </div>
                                )}

                                {/* Mail Access Details for all email operations */}
                                {emailOps.includes(entry.Operation) && (
                                  <div>
                                    <div className="mb-2 flex items-center gap-2 text-blue-700 dark:text-blue-300">
                                      <Mail className="h-5 w-5" />
                                      <span className="font-semibold">Mail Access Details</span>
                                    </div>
                                    <div className="bg-blue-50 dark:bg-blue-900/20 border border-blue-200 dark:border-blue-800 rounded-lg p-4">
                                      <div className="space-y-3">
                                        {(() => {
                                          let parsedAuditData: AuditData | null = null;
                                          try {
                                            parsedAuditData = typeof entry.AuditData === 'string' 
                                              ? JSON.parse(entry.AuditData)
                                              : entry.AuditData;
                                          } catch (e) {
                                            console.warn('Failed to parse AuditData:', e);
                                          }

                                          if (!parsedAuditData) return null;

                                          return (
                                            <>
                                              {parsedAuditData.MailboxOwnerUPN && (
                                                <div className="flex items-start gap-3 p-2 bg-white/50 dark:bg-black/10 rounded border border-blue-100 dark:border-blue-900">
                                                  <User className="h-5 w-5 text-blue-600 dark:text-blue-400 flex-shrink-0 mt-0.5" />
                                                  <div>
                                                    <div className="text-sm font-medium text-blue-700 dark:text-blue-300">Mailbox Owner</div>
                                                    <div className="text-sm text-slate-900 dark:text-slate-100">{parsedAuditData.MailboxOwnerUPN}</div>
                                                    </div>
                                                </div>
                                              )}
                                              {parsedAuditData.ClientInfoString && (
                                                <div className="flex items-start gap-3 p-2 bg-white/50 dark:bg-black/10 rounded border border-blue-100 dark:border-blue-900">
                                                  <Monitor className="h-5 w-5 text-blue-600 dark:text-blue-400 flex-shrink-0 mt-0.5" />
                                                  <div>
                                                    <div className="text-sm font-medium text-blue-700 dark:text-blue-300">Client Info</div>
                                                    <div className="text-sm text-slate-900 dark:text-slate-100 break-all">{parsedAuditData.ClientInfoString}</div>
                                                  </div>
                                                </div>
                                              )}
                                              {parsedAuditData.OperationProperties?.find(prop => prop.Name === "MailAccessType") && (
                                                <div className="flex items-start gap-3 p-2 bg-white/50 dark:bg-black/10 rounded border border-blue-100 dark:border-blue-900">
                                                  <Network className="h-5 w-5 text-blue-600 dark:text-blue-400 flex-shrink-0 mt-0.5" />
                                                  <div>
                                                    <div className="text-sm font-medium text-blue-700 dark:text-blue-300">Mail Access Type</div>
                                                    <div className="text-sm text-slate-900 dark:text-slate-100">
                                                      {parsedAuditData.OperationProperties.find(prop => prop.Name === "MailAccessType")?.Value}
                                                    </div>
                                                  </div>
                                                </div>
                                              )}
                                              {entry.Operation === "Send" && parsedAuditData.Subject && (
                                                <div className="flex items-start gap-3 p-2 bg-white/50 dark:bg-black/10 rounded border border-blue-100 dark:border-blue-900">
                                                  <Mail className="h-5 w-5 text-blue-600 dark:text-blue-400 flex-shrink-0 mt-0.5" />
                                                  <div>
                                                    <div className="text-sm font-medium text-blue-700 dark:text-blue-300">Email Subject</div>
                                                    <div className="text-sm text-slate-900 dark:text-slate-100">{parsedAuditData.Subject}</div>
                                                  </div>
                                                </div>
                                              )}
                                            </>
                                          );
                                        })()}
                                      </div>
                                    </div>
                                  </div>
                                )}

                                {entry.UserAgentInfo && (
                                  <div>
                                    <div className="mb-2 flex items-center gap-2 text-blue-700 dark:text-blue-300">
                                      <Monitor className="h-5 w-5" />
                                      <span className="font-semibold">Client Details</span>
                                    </div>
                                    <div className="bg-blue-50 dark:bg-blue-900/20 border border-blue-200 dark:border-blue-800 rounded-lg p-4">
                                      <div className="space-y-3">
                                        <div className="flex items-start gap-3 p-2 bg-white/50 dark:bg-black/10 rounded border border-blue-100 dark:border-blue-900">
                                          <Globe className="h-5 w-5 text-blue-600 dark:text-blue-400 flex-shrink-0 mt-0.5" />
                                          <div>
                                            <div className="text-sm font-medium text-blue-700 dark:text-blue-300">
                                              Browser
                                            </div>
                                            <div className="mt-1 text-sm text-slate-900 dark:text-slate-100">
                                              {entry.UserAgentInfo.browser}
                                              {entry.UserAgentInfo.browserVersion !== "Unknown" && (
                                                <span className="text-slate-500 dark:text-slate-400"> v{entry.UserAgentInfo.browserVersion}</span>
                                              )}
                                            </div>
                                          </div>
                                        </div>

                                        <div className="flex items-start gap-3 p-2 bg-white/50 dark:bg-black/10 rounded border border-blue-100 dark:border-blue-900">
                                          <Monitor className="h-5 w-5 text-blue-600 dark:text-blue-400 flex-shrink-0 mt-0.5" />
                                          <div>
                                            <div className="text-sm font-medium text-blue-700 dark:text-blue-300">
                                              Operating System
                                            </div>
                                            <div className="mt-1 text-sm text-slate-900 dark:text-slate-100">
                                              {entry.UserAgentInfo.os}
                                              {entry.UserAgentInfo.osVersion !== "Unknown" && (
                                                <span className="text-slate-500 dark:text-slate-400"> {entry.UserAgentInfo.osVersion}</span>
                                              )}
                                            </div>
                                          </div>
                                        </div>

                                        <div className="flex items-start gap-3 p-2 bg-white/50 dark:bg-black/10 rounded border border-blue-100 dark:border-blue-900">
                                          <Laptop className="h-5 w-5 text-blue-600 dark:text-blue-400 flex-shrink-0 mt-0.5" />
                                          <div>
                                            <div className="text-sm font-medium text-blue-700 dark:text-blue-300">
                                              Device Type
                                            </div>
                                            <div className="mt-1 text-sm text-slate-900 dark:text-slate-100">
                                              {entry.UserAgentInfo.device}
                                              {entry.UserAgentInfo.isMobile && (
                                                <span className="text-slate-500 dark:text-slate-400"> (Mobile Device)</span>
                                              )}
                                            </div>
                                          </div>
                                        </div>

                                        <div className="flex items-start gap-3 p-2 bg-white/50 dark:bg-black/10 rounded border border-blue-100 dark:border-blue-900">
                                          <Code className="h-5 w-5 text-blue-600 dark:text-blue-400 flex-shrink-0 mt-0.5" />
                                          <div>
                                            <div className="text-sm font-medium text-blue-700 dark:text-blue-300">
                                              Raw User Agent
                                            </div>
                                            <div className="mt-1 text-xs text-slate-500 dark:text-slate-400 break-all">
                                              {entry.UserAgentInfo.raw}
                                            </div>
                                          </div>
                                        </div>
                                      </div>
                                    </div>
                                  </div>
                                )}

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

                                {entry.Operation === "Update application – Certificates and secrets management " && entry.ModifiedProperties && (
                                <div>
                                    <div className="mb-2 flex items-center gap-2 text-red-700 dark:text-red-300">
                                      <Shield className="h-5 w-5" />
                                      <span className="font-semibold">Certificate & Secret Changes</span>
                                    </div>
                                    <div className="bg-red-50 dark:bg-red-900/20 border border-red-200 dark:border-red-800 rounded-lg p-4">
                                      <div className="space-y-3">
                                        {(() => {
                                          let modifiedProps: any[] = [];
                                          try {
                                            if (typeof entry.ModifiedProperties === 'string') {
                                              modifiedProps = JSON.parse(entry.ModifiedProperties);
                                            } else {
                                              modifiedProps = entry.ModifiedProperties;
                                            }
                                          } catch (e) {
                                            console.warn('Failed to parse ModifiedProperties:', e);
                                            return null;
                                          }

                                          const keyDescriptionProp = modifiedProps.find(prop => prop.Name === 'KeyDescription');
                                          if (!keyDescriptionProp) {
                                            return (
                                              <div className="text-sm text-slate-600 dark:text-slate-400">
                                                No certificate or secret changes found
                                              </div>
                                            );
                                          }

                                          // Parse the old and new values
                                          const parseKeyList = (value: string) => {
                                            try {
                                              const parsed = JSON.parse(value);
                                              return parsed.map((item: string) => {
                                                const match = item.match(/KeyIdentifier=([^,]+),KeyType=([^,]+),KeyUsage=([^,]+),DisplayName=([^\]]+)/);
                                                if (match) {
                                                  return {
                                                    keyId: match[1],
                                                    keyType: match[2],
                                                    keyUsage: match[3],
                                                    displayName: match[4]
                                                  };
                                                }
                                                return null;
                                              }).filter(Boolean);
                                            } catch (e) {
                                              console.warn('Failed to parse key list:', e);
                                              return [];
                                            }
                                          };

                                          const oldKeys = parseKeyList(keyDescriptionProp.OldValue);
                                          const newKeys = parseKeyList(keyDescriptionProp.NewValue);

                                          // Find added and removed keys
                                          const addedKeys = newKeys.filter(newKey => 
                                            !oldKeys.some(oldKey => oldKey.keyId === newKey.keyId)
                                          );
                                          const removedKeys = oldKeys.filter(oldKey => 
                                            !newKeys.some(newKey => newKey.keyId === oldKey.keyId)
                                          );

                                          return (
                                            <>
                                              {addedKeys.length > 0 && (
                                                <div className="space-y-2">
                                                  <div className="text-sm font-medium text-green-600 dark:text-green-400">Added:</div>
                                                  {addedKeys.map((key, index) => (
                                                    <div key={index} className="flex items-start gap-3 p-2 bg-white/50 dark:bg-black/10 rounded border border-green-100 dark:border-green-900">
                                                      <Shield className="h-5 w-5 text-green-600 dark:text-green-400 flex-shrink-0 mt-0.5" />
                                                      <div>
                                                        <div className="text-sm font-medium text-green-700 dark:text-green-300">
                                                          {key.displayName}
                                                        </div>
                                                        <div className="mt-1 text-xs text-slate-500 dark:text-slate-400">
                                                          Type: {key.keyType}, Usage: {key.keyUsage}
                                                        </div>
                                                      </div>
                                                    </div>
                                                  ))}
                                                </div>
                                              )}
                                              {removedKeys.length > 0 && (
                                                <div className="space-y-2">
                                                  <div className="text-sm font-medium text-red-600 dark:text-red-400">Removed:</div>
                                                  {removedKeys.map((key, index) => (
                                                    <div key={index} className="flex items-start gap-3 p-2 bg-white/50 dark:bg-black/10 rounded border border-red-100 dark:border-red-900">
                                                      <Shield className="h-5 w-5 text-red-600 dark:text-red-400 flex-shrink-0 mt-0.5" />
                                                      <div>
                                                        <div className="text-sm font-medium text-red-700 dark:text-red-300">
                                                          {key.displayName}
                                                        </div>
                                                        <div className="mt-1 text-xs text-slate-500 dark:text-slate-400">
                                                          Type: {key.keyType}, Usage: {key.keyUsage}
                                                        </div>
                                                      </div>
                                                    </div>
                                                  ))}
                                                </div>
                                              )}
                                              {addedKeys.length === 0 && removedKeys.length === 0 && (
                                                <div className="text-sm text-slate-600 dark:text-slate-400">
                                                  Keys were reordered but no additions or removals were made
                                                </div>
                                              )}
                                            </>
                                          );
                                        })()}
                                      </div>
                                    </div>
                                  </div>
                                )}

                                <div>
                                  <div className="font-medium mb-1">Raw Modified Properties</div>
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

      {/* IP Export Confirmation Modal */}
      {showIPExportModal && (
        <div className="fixed inset-0 bg-black/50 flex items-center justify-center z-50">
          <div className="bg-white dark:bg-gray-800 p-6 rounded-lg shadow-lg max-w-md w-full">
            <h3 className="text-lg font-semibold mb-4">Confirm IP Export</h3>
            <p className="text-sm text-gray-600 dark:text-gray-300 mb-4">
              You are about to send IP addresses to IPinfo.io for visualization. This will:
            </p>
            <ul className="list-disc list-inside text-sm text-gray-600 dark:text-gray-300 mb-6 space-y-2">
              <li>Send unique IP addresses from login events</li>
              <li>Create a visualization map at IPinfo.io</li>
              <li>Open the map in a new tab</li>
            </ul>
            <div className="flex justify-end gap-3">
              <button
                onClick={() => setShowIPExportModal(false)}
                className="px-4 py-2 text-sm text-gray-600 hover:text-gray-800 dark:text-gray-300 dark:hover:text-white"
              >
                Cancel
              </button>
              <button
                onClick={() => {
                  setShowIPExportModal(false);
                  exportIPsToMap();
                }}
                className="px-4 py-2 text-sm bg-blue-600 hover:bg-blue-700 text-white rounded"
              >
                Proceed
              </button>
            </div>
          </div>
        </div>
      )}

      <style jsx>{`
        @keyframes fadeInUp {
          from {
            opacity: 0;
            transform: translateY(10px);
          }
          to {
            opacity: 1;
            transform: translateY(0);
          }
        }
        
        .animate-fade-in-up {
          animation: fadeInUp 0.3s ease-out;
        }
      `}</style>
    </div>
  )
}

