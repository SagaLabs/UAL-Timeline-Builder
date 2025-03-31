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
  Code
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
  CorrelationId?: string;
  ModifiedProperties?: string;
  AuditDataRaw?: string;
  UserAgent?: string;
  UserAgentInfo?: UserAgentInfo;
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

export default function UALTimelineBuilder() {
  const [logs, setLogs] = useState<LogEntry[]>([])
  const [fileNames, setFileNames] = useState<string[]>([])
  const [userFilters, setUserFilters] = useState<string[]>([])
  const [workloadFilters, setWorkloadFilters] = useState<string[]>([])
  const [operationFilters, setOperationFilters] = useState<string[]>([])
  const [correlationFilter, setCorrelationFilter] = useState("")
  const [showOnlyRisky, setShowOnlyRisky] = useState(false)
  const [visibleCount, setVisibleCount] = useState(100)
  const [loading, setLoading] = useState(false)
  const [searchTerm, setSearchTerm] = useState("")
  const [darkMode, setDarkMode, mounted] = useDarkMode()

  const [userDropdownOpen, setUserDropdownOpen] = useState(false)
  const [workloadDropdownOpen, setWorkloadDropdownOpen] = useState(false)
  const [operationDropdownOpen, setOperationDropdownOpen] = useState(false)
  const [userSearchTerm, setUserSearchTerm] = useState("")
  const [workloadSearchTerm, setWorkloadSearchTerm] = useState("")
  const [operationSearchTerm, setOperationSearchTerm] = useState("")
  const [showIPExportModal, setShowIPExportModal] = useState(false)

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
          if (entry.Operation === "UpdateInboxRule" || entry.Operation === "New-InboxRule") {
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

  const filteredLogs = useMemo(() => {
    return logs.filter((entry, index) => {
      // Multi-select filters - if no filters selected, show all
      const userMatch = userFilters.length === 0 || userFilters.includes(entry.UserId || entry.UserKey || '')
      const workloadMatch = workloadFilters.length === 0 || workloadFilters.includes(entry.Workload || '')
      const operationMatch = operationFilters.length === 0 || operationFilters.includes(entry.Operation || '')

      // Single-select filters
      const correlationMatch = correlationFilter ? entry.CorrelationId === correlationFilter : true
      const riskyMatch = showOnlyRisky ? riskyOps.includes(entry.Operation || '') : true

      // Optimized search functionality
      const searchMatch = !searchTerm || (() => {
        const searchLower = searchTerm.toLowerCase();
        const searchableEntry = searchableFields[index];
        
        return (
          searchableEntry.id.toLowerCase().includes(searchLower) ||
          searchableEntry.operation.toLowerCase().includes(searchLower) ||
          searchableEntry.workload.toLowerCase().includes(searchLower) ||
          searchableEntry.subject.toLowerCase().includes(searchLower) ||
          searchableEntry.messageId.toLowerCase().includes(searchLower) ||
          searchableEntry.correlationId.toLowerCase().includes(searchLower) ||
          searchableEntry.clientIP.toLowerCase().includes(searchLower) ||
          searchableEntry.fileName.toLowerCase().includes(searchLower) ||
          searchableEntry.userAgent.toLowerCase().includes(searchLower)
        );
      })();

      return userMatch && workloadMatch && operationMatch && correlationMatch && riskyMatch && searchMatch
    })
  }, [logs, userFilters, workloadFilters, operationFilters, correlationFilter, showOnlyRisky, searchTerm, searchableFields])

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

  const downloadInternetMessageIds = () => {
    // Create a map to store message IDs and their details
    const messageStats = new Map<string, {
      readBy: Set<string>;
      readAt: string[];
      subject: string;
      workload: string;
      clientIP: string;
      folderPath: string;
      sizeInBytes: number;
    }>();

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
                  const existing = messageStats.get(item.InternetMessageId) || {
                    readBy: new Set<string>(),
                    readAt: [],
                    subject: entry.Subject || 'N/A',
                    workload: entry.Workload || 'Unknown',
                    clientIP: entry.ClientIP || 'N/A',
                    folderPath: folder.Path || 'Unknown',
                    sizeInBytes: item.SizeInBytes || 0
                  };

                  // Add user who read the message
                  if (entry.UserId || entry.UserKey) {
                    existing.readBy.add(entry.UserId || entry.UserKey);
                  }

                  // Add read timestamp
                  if (entry.TimeGenerated) {
                    existing.readAt.push(entry.TimeGenerated);
                  }

                  messageStats.set(item.InternetMessageId, existing);
                }
              });
            }
          });
        }
      } catch (e) {
        console.warn("Error processing log entry:", e);
      }
    });

    // Convert to CSV format
    const csvRows = ['InternetMessageId,Subject,Workload,Read By,Read Timestamps,Client IP,Folder Path,Size (Bytes)'];
    messageStats.forEach((data, messageId) => {
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
    a.download = 'internet-message-ids.csv';
    a.click();
    URL.revokeObjectURL(url);
  };

  const downloadIPStats = () => {
    // Define login operations
    const loginOperations = [
      "UserLoggedIn",
      "SignIn",
      "UserLoginFailed",
      "UserLoginSuccess"
    ];

    // Create a map to store IP statistics
    const ipStats = new Map<string, {
      users: Set<string>;
      timestamps: string[];
      operations: Set<string>;
      workloads: Set<string>;
    }>();

    // Process each log entry
    logs.forEach(entry => {
      // Only process login operations
      if (!loginOperations.includes(entry.Operation)) return;

      // Skip if user filters are active and this entry doesn't match any selected user
      if (userFilters.length > 0) {
        const entryUser = entry.UserId || entry.UserKey;
        if (!entryUser || !userFilters.includes(entryUser)) return;
      }

      const ip = entry.ClientIP;
      if (!ip) return;

      const existing = ipStats.get(ip) || {
        users: new Set<string>(),
        timestamps: [],
        operations: new Set<string>(),
        workloads: new Set<string>()
      };

      // Add user
      if (entry.UserId || entry.UserKey) {
        existing.users.add(entry.UserId || entry.UserKey);
      }

      // Add timestamp
      if (entry.TimeGenerated) {
        existing.timestamps.push(entry.TimeGenerated);
      }

      // Add operation
      if (entry.Operation) {
        existing.operations.add(entry.Operation);
      }

      // Add workload
      if (entry.Workload) {
        existing.workloads.add(entry.Workload);
      }

      ipStats.set(ip, existing);
    });

    // Convert to CSV format
    const csvRows = ['IP Address,Users,First Seen,Last Seen,Operations,Workloads'];
    ipStats.forEach((data, ip) => {
      const users = Array.from(data.users).join('; ');
      const timestamps = data.timestamps.sort();
      const firstSeen = timestamps[0] || 'N/A';
      const lastSeen = timestamps[timestamps.length - 1] || 'N/A';
      const operations = Array.from(data.operations).join('; ');
      const workloads = Array.from(data.workloads).join('; ');
      
      csvRows.push(`"${ip}","${users}","${firstSeen}","${lastSeen}","${operations}","${workloads}"`);
    });

    // Create and download the file
    const csvContent = csvRows.join('\n');
    const blob = new Blob([csvContent], { type: 'text/csv;charset=utf-8;' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = 'ip-login-stats.csv';
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

  return (
    <div className="min-h-screen bg-white dark:bg-gray-900 p-4 md:p-6">
      <div className="max-w-6xl mx-auto bg-white dark:bg-gray-900 shadow-xl rounded-xl border border-slate-200 dark:border-gray-700 overflow-hidden">
        {/* Header */}
        <div className="bg-gradient-to-r from-blue-600 to-indigo-600 p-6 text-white">
          <div className="flex items-center justify-between">
          <h1 className="text-3xl font-bold flex items-center gap-2">
            <Shield className="h-8 w-8" />
            UAL-Timeline-Builder (UTB)
          </h1>
            {renderDarkModeButton()}
          </div>
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
                    <div className="relative flex-1">
                      <Search className="absolute left-3 top-1/2 transform -translate-y-1/2 h-4 w-4 text-slate-400" />
                      <input
                        type="text"
                        placeholder="Search logs..."
                        onChange={(e) => debouncedSetSearchTerm(e.target.value)}
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
                      onClick={downloadInternetMessageIds}
                      className="flex items-center justify-center gap-1.5 px-3 py-1.5 bg-slate-800 hover:bg-slate-900 dark:bg-slate-700 dark:hover:bg-gray-600 text-white rounded-lg transition-colors text-sm"
                    >
                      <Mail className="h-3.5 w-3.5" />
                      <span>Export Mail Activity</span>
                      <span className="text-xs text-slate-300 dark:text-slate-400">{userFilters.length > 0 ? `${userFilters.length} users` : 'all users'}</span>
                    </button>
                    <button
                      onClick={downloadIPStats}
                      className="flex items-center justify-center gap-1.5 px-3 py-1.5 bg-slate-800 hover:bg-slate-900 dark:bg-slate-700 dark:hover:bg-gray-600 text-white rounded-lg transition-colors text-sm"
                    >
                      <Globe className="h-3.5 w-3.5" />
                      <span>Export IP Stats</span>
                      <span className="text-xs text-slate-300 dark:text-slate-400">{userFilters.length > 0 ? `${userFilters.length} users` : 'all users'}</span>
                    </button>
                    <button
                      onClick={() => setShowIPExportModal(true)}
                      className="flex items-center justify-center gap-1.5 px-3 py-1.5 bg-blue-600 hover:bg-blue-700 text-white rounded-lg transition-colors text-sm"
                    >
                      <MapPin className="h-3.5 w-3.5" />
                      <span>Map IPs</span>
                      <span className="text-xs text-blue-200">{userFilters.length > 0 ? `${userFilters.length} users` : 'all users'}</span>
                    </button>
                    <button
                      onClick={downloadNDJSON}
                      className="flex items-center justify-center gap-1.5 px-3 py-1.5 bg-slate-800 hover:bg-slate-900 dark:bg-slate-700 dark:hover:bg-gray-600 text-white rounded-lg transition-colors text-sm"
                    >
                      <Download className="h-3.5 w-3.5" />
                      <span>Export to NDJSON</span>
                      <span className="text-xs text-slate-300 dark:text-slate-400">{userFilters.length > 0 ? `${userFilters.length} users` : 'all users'}</span>
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
                            <div className="flex items-center justify-between mb-2">
                            <button
                              onClick={() => setUserFilters([])}
                              className="text-xs text-blue-600 dark:text-blue-400 hover:underline"
                            >
                              Clear all
                            </button>
                              <div className="relative">
                                <Search className="absolute left-2 top-1/2 transform -translate-y-1/2 h-3 w-3 text-slate-400" />
                                <input
                                  type="text"
                                  placeholder="Search users..."
                                  value={userSearchTerm}
                                  onChange={(e) => setUserSearchTerm(e.target.value)}
                                  className="pl-8 pr-2 py-1 text-xs w-full rounded border border-slate-300 dark:border-slate-600 bg-white dark:bg-gray-800 focus:ring-1 focus:ring-blue-500 focus:border-blue-500 dark:focus:ring-blue-600 dark:focus:border-blue-600 outline-none"
                                />
                              </div>
                            </div>
                          </div>
                          <div className="p-1">
                            {filteredUserOptions.map((user) => (
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
                            <div className="flex items-center justify-between mb-2">
                            <button
                              onClick={() => setWorkloadFilters([])}
                              className="text-xs text-blue-600 dark:text-blue-400 hover:underline"
                            >
                              Clear all
                            </button>
                              <div className="relative">
                                <Search className="absolute left-2 top-1/2 transform -translate-y-1/2 h-3 w-3 text-slate-400" />
                                <input
                                  type="text"
                                  placeholder="Search workloads..."
                                  value={workloadSearchTerm}
                                  onChange={(e) => setWorkloadSearchTerm(e.target.value)}
                                  className="pl-8 pr-2 py-1 text-xs w-full rounded border border-slate-300 dark:border-slate-600 bg-white dark:bg-gray-800 focus:ring-1 focus:ring-blue-500 focus:border-blue-500 dark:focus:ring-blue-600 dark:focus:border-blue-600 outline-none"
                                />
                              </div>
                            </div>
                          </div>
                          <div className="p-1">
                            {filteredWorkloadOptions.map((workload) => (
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
                            <div className="flex items-center justify-between mb-2">
                            <button
                              onClick={() => setOperationFilters([])}
                              className="text-xs text-blue-600 dark:text-blue-400 hover:underline"
                            >
                              Clear all
                            </button>
                              <div className="relative">
                                <Search className="absolute left-2 top-1/2 transform -translate-y-1/2 h-3 w-3 text-slate-400" />
                                <input
                                  type="text"
                                  placeholder="Search operations..."
                                  value={operationSearchTerm}
                                  onChange={(e) => setOperationSearchTerm(e.target.value)}
                                  className="pl-8 pr-2 py-1 text-xs w-full rounded border border-slate-300 dark:border-slate-600 bg-white dark:bg-gray-800 focus:ring-1 focus:ring-blue-500 focus:border-blue-500 dark:focus:ring-blue-600 dark:focus:border-blue-600 outline-none"
                                />
                              </div>
                            </div>
                          </div>
                          <div className="p-1">
                            {filteredOperationOptions.map((operation) => (
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
  {/* Special display for inbox rules */}
  {(entry.Operation === "UpdateInboxRule" || entry.Operation === "New-InboxRule") &&
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
    </div>
  )
}

