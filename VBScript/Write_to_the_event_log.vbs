' Constants for type of event log entry
const EVENTLOG_SUCCESS = 0
const EVENTLOG_ERROR = 1
const EVENTLOG_WARNING = 2
const EVENTLOG_INFORMATION = 4
const EVENTLOG_AUDIT_SUCCESS = 8
const EVENTLOG_AUDIT_FAILURE = 16

strMessage = "My event log message..."

set objShell = CreateObject("WScript.Shell")
objShell.LogEvent EVENTLOG_INFORMATION, strMessage