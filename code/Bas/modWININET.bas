Attribute VB_Name = "modWININET"
Option Explicit


'=============================================================================================================
'
' modWININET Module
' -----------------
''
' VB Versions : 5.0 / 6.0
'
' Requires    : At least IE 3.0 or an operating system that comes with at least IE 3.0 pre-installed.  Some
'               WININET.DLL APIs require that IE 4.0 or 5.0 be installed.  See below for details.
'
' Description : This module gives full access to all the documented functions of the WININET.DLL and all
'               of the types (structures) and constants that are required to make use of those functions.
'
' See Also:
' ---------
' http://msdn.microsoft.com/workshop/networking/wininet/reference/functions/all_functions.asp
' http://msdn.microsoft.com/workshop/networking/wininet/overview/appendix_a.asp
' http://msdn.microsoft.com/workshop/networking/wininet/overview/ftp.asp
' http://msdn.microsoft.com/workshop/networking/wininet/overview/http.asp
' http://msdn.microsoft.com/workshop/networking/wininet/overview/gopher.asp
' http://msdn.microsoft.com/workshop/networking/wininet/overview/introduction.asp
' http://www.vbip.com/default.asp
'
'=============================================================================================================
'
'¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯
' The following table is an alphabetical list of the functions provided by the Microsoft® Win32® Internet API.
'¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯
' CommitUrlCacheEntry             Stores data in the specified file in the Internet cache and associates it with the given URL.
' CreateUrlCacheEntry             Creates a local file name for saving the cache entry based on the specified URL and the file extension.
' CreateUrlCacheGroup             Generates cache group identifications.
' DeleteUrlCacheEntry             Removes the file associated with the source name from the cache, if the file exists.
' DeleteUrlCacheGroup             Releases the specified GROUPID and any associated state in the cache index file.
' FindCloseUrlCache               Closes the specified cache enumeration handle.
' FindFirstUrlCacheEntry          Begins the enumeration of the Internet cache.
' FindFirstUrlCacheEntryEx        Starts a filtered enumeration of the Internet cache.
' FindFirstUrlCacheGroup          Initiates the enumeration of the cache groups in the Internet cache.
' FindNextUrlCacheEntry           Retrieves the next entry in the Internet cache.
' FindNextUrlCacheEntryEx         Finds the next cache entry in a cache enumeration started by the FindFirstUrlCacheEntryEx function.
' FindNextUrlCacheGroup           Retrieves the next cache group in a cache group enumeration started by FindFirstUrlCacheGroup.
' FtpCommand                      Allows an application to send commands directly to an FTP server.
' FtpCreateDirectory              Creates a new directory on the FTP server.
' FtpDeleteFile                   Deletes a file stored on the FTP server.
' FtpFindFirstFile                Searches the specified directory of the given FTP session. File and directory entries are returned to the application in the WIN32_FIND_DATA structure.
' FtpGetCurrentDirectory          Retrieves the current directory for the specified FTP session.
' FtpGetFile                      Retrieves a file from the FTP server and stores it under the specified file name, creating a new local file in the process.
' FtpGetFileSize                  Retrieves the file size of the requested FTP resource.
' FtpOpenFile                     Initiates access to a remote file on an FTP server for reading or writing.
' FtpPutFile                      Stores a file on the FTP server.
' FtpRemoveDirectory              Removes the specified directory on the FTP server.
' FtpRenameFile                   Renames a file stored on the FTP server.
' FtpSetCurrentDirectory          Changes to a different working directory on the FTP server.
' GetUrlCacheEntryInfo            Retrieves information about a cache entry.
' GetUrlCacheEntryInfoEx          Searches for the URL after translating any cached redirections that would be applied in offline mode by HttpSendRequest.
' GetUrlCacheGroupAttribute       Retrieves the attribute information of the specified cache group.
' GopherCreateLocator             Creates a Gopher or Gopher+ locator string from its component parts.
' GopherFindFirstFile             Uses a Gopher locator and some search criteria to create a session with the server and locate the requested documents, binary files, index servers, or directory trees.
' GopherGetAttribute              Retrieves the specific attribute information from the server.
' GopherGetLocatorType            Parses a Gopher locator and determines its attributes.
' GopherOpenFile                  Begins reading a Gopher data file from a Gopher server.
' HttpAddRequestHeaders           Adds one or more HTTP request headers to the HTTP request handle.
' HttpEndRequest                  Ends an HTTP request that was initiated by HttpSendRequestEx.
' HttpOpenRequest                 Creates an HTTP request handle.
' HttpQueryInfo                   Retrieves header information associated with an HTTP request.
' HttpSendRequest                 Sends the specified request to the HTTP server.
' HttpSendRequestEx               Sends the specified request to the HTTP server and allows chunked transfers.
' InternetAttemptConnect          Attempts to make a connection to the Internet.
' InternetAutodial                Causes the modem to automatically dial the default Internet connection.
' InternetAutodialHangup          Disconnects an automatic dial-up connection.
' InternetCanonicalizeUrl         Canonicalizes a URL, which includes converting unsafe characters and spaces into escape sequences.
' InternetCheckConnection         Allows an application to check if a connection to the Internet can be established.
' InternetCloseHandle             Closes a single Internet handle.
' InternetCombineUrl              Combines a base and relative URL into a single URL. The resultant URL will be canonicalized (see InternetCanonicalizeUrl).
' InternetConfirmZoneCrossing     Checks for changes between secure and nonsecure URLs. When a change occurs in security between two URLs, an application should allow the user to acknowledge this change, typically by displaying a dialog box.
' InternetConnect                 Opens an FTP, Gopher, or HTTP session for a given site.
' InternetCrackUrl                Cracks a URL into its component parts.
' InternetCreateUrl               Creates a URL from its component parts.
' InternetDial                    Initiates a connection to the Internet using a modem.
' InternetErrorDlg                Displays a dialog box for the error that is passed to InternetErrorDlg, if an appropriate dialog box exists. If the FLAGS_ERROR_UI_FILTER_FOR_ERRORS flag is used, the function also checks the headers for any hidden errors and displays a dialog box if needed.
' InternetFindNextFile            Continues a file search started as a result of a previous call to FtpFindFirstFile or GopherFindFirstFile.
' InternetGetConnectedState       Retrieves the connected state of the local system.
' InternetGetConnectedStateEx     Retrieves the connected state of the specified Internet connection.
' InternetGetCookie               Retrieves the cookie for the specified URL.
' InternetGetLastResponseInfo     Retrieves the last Win32® Internet function error description or server response on the thread calling this function.
' InternetGoOnline                Prompts the user for permission to initiate connection to a URL.
' InternetHangUp                  Instructs the modem to disconnect from the Internet.
' InternetInitializeAutoProxyDll  Not currently supported.
' InternetLockRequestFile         Allows the user to place a lock on the file that is being used.
' InternetOpen                    Initializes an application's use of the Win32® Internet functions.
' InternetOpenUrl                 Begins reading a complete FTP, Gopher, or HTTP URL. Use InternetCanonicalizeUrl first if the URL being used contains a relative URL and a base URL separated by blank spaces.
' InternetQueryDataAvailable      Queries the server to determine the amount of data available.
' InternetQueryOption             Queries an Internet option on the specified handle.
' InternetReadFile                Reads data from a handle opened by the InternetOpenUrl, FtpOpenFile, GopherOpenFile, or HttpOpenRequest function.
' InternetReadFileEx              Reads data from a handle opened by the InternetOpenUrl, FtpOpenFile, GopherOpenFile, or HttpOpenRequest function.
' InternetSetCookie               Creates a cookie associated with the specified URL.
' InternetSetDialState            Obsolete. Do not use.
' InternetSetFilePointer          Sets a file position for InternetReadFile. This is a synchronous call; however, subsequent calls to InternetReadFile might block or return pending if the data is not available from the cache and the server does not support random access.
' InternetSetOption               Sets an Internet option.
' InternetSetOptionEx             Not currently implemented.
' InternetSetStatusCallback       Sets up a callback function that Win32® Internet functions can call as progress is made during an operation.
' InternetTimeFromSystemTime      Formats a date and time according to the HTTP version 1.0 specification.
' InternetTimeToSystemTime        Takes an HTTP time/date string and converts it to a SYSTEMTIME structure.
' InternetUnlockRequestFile       Unlocks a file that was locked using InternetLockRequestFile.
' InternetWriteFile               Writes data to an open Internet file.
' ReadUrlCacheEntryStream         Reads the cached data from a stream that has been opened using the RetrieveUrlCacheEntryStream function.
' RetrieveUrlCacheEntryFile       Locks the cache entry file associated with the specified URL.
' RetrieveUrlCacheEntryStream     Provides the most efficient and implementation-independent way of accessing the cache data.
' SetUrlCacheEntryGroup           Adds entries to or removes entries from a cache group.
' SetUrlCacheEntryInfo            Sets the specified members of the INTERNET_CACHE_ENTRY_INFO structure.
' SetUrlCacheGroupAttribute       Sets the attribute information of the specified cache group.
' UnlockUrlCacheEntryFile         Unlocks the cache entry that was locked while the file was retrieved for use from the cache.
' UnlockUrlCacheEntryStream       Closes the stream that has been retrieved using the RetrieveUrlCacheEntryStream function.
'_____________________________________________________________________________________________________________
'¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯




'_____________________________________________________________________________________________________________
'¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯
' The following functions are only available if you have Internet Explorer 3.0 or greater installed on your
' computer, or one of the following operating systems (which have at least IE 3.0 pre-installed):
'¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯
'   - Windows 95B
'   - Windows 98 (First Edition)
'   - Windows 98 (Second Edition)
'   - Windows NT 4.0
'   - Windows 2000
'   - Windows ME
'¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯
'  CommitUrlCacheEntry
'  CreateUrlCacheEntry
'  DeleteUrlCacheEntry
'  FindCloseUrlCache
'  FindFirstUrlCacheEntry
'  FindNextUrlCacheEntry
'  FtpCreateDirectory
'  FtpDeleteFile
'  FtpFindFirstFile
'  FtpGetCurrentDirectory
'  FtpGetFile
'  FtpOpenFile
'  FtpPutFile
'  FtpRemoveDirectory
'  FtpRenameFile
'  FtpSetCurrentDirectory
'  GetUrlCacheEntryInfo
'  GopherCreateLocator
'  GopherFindFirstFile
'  GopherGetAttribute
'  GopherGetLocatorType
'  GopherOpenFile
'  HttpAddRequestHeaders
'  HttpOpenRequest
'  HttpQueryInfo
'  HttpSendRequest
'  InternetAttemptConnect
'  InternetCanonicalizeUrl
'  InternetCheckConnection
'  InternetCloseHandle
'  InternetCombineUrl
'  InternetConfirmZoneCrossing
'  InternetConnect
'  InternetCrackUrl
'  InternetCreateUrl
'  InternetErrorDlg
'  InternetFindNextFile
'  InternetGetCookie
'  InternetGetLastResponseInfo
'  InternetLockRequestFile
'  InternetOpen
'  InternetOpenUrl
'  InternetQueryDataAvailable
'  InternetQueryOption
'  InternetReadFile
'  InternetSetCookie
'  InternetSetFilePointer
'  InternetSetOption
'  InternetSetStatusCallback
'  InternetTimeFromSystemTime
'  InternetTimeToSystemTime
'  InternetUnlockRequestFile
'  InternetWriteFile
'  ReadUrlCacheEntryStream
'  RetrieveUrlCacheEntryFile
'  RetrieveUrlCacheEntryStream
'  SetUrlCacheEntryInfo
'  UnlockUrlCacheEntryFile
'  UnlockUrlCacheEntryStream
'
'_____________________________________________________________________________________________________________
' The following functions are only available if you have Internet Explorer 4.0 or greater installed on your
' computer, or one of the following operating systems (which have at least IE 4.0 pre-installed):
'¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯
'   - Windows 98 (First Edition)
'   - Windows 98 (Second Edition)
'   - Windows 2000
'   - Windows ME
'¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯
'  CreateUrlCacheGroup
'  DeleteUrlCacheGroup
'  FindFirstUrlCacheEntryEx
'  FindNextUrlCacheEntryEx
'  GetUrlCacheEntryInfoEx
'  HttpEndRequest
'  HttpSendRequestEx
'  InternetAutodial
'  InternetAutodialHangup
'  InternetDial
'  InternetGetConnectedState
'  InternetGoOnline
'  InternetHangUp
'  InternetReadFileEx
'  SetUrlCacheEntryGroup
'
'_____________________________________________________________________________________________________________
' The following functions are only available if you have Internet Explorer 5.0 or greater installed on your
' computer, or one of the following operating systems (which have at least IE 5.0 pre-installed):
'¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯
'   - Windows 98 (Second Edition)
'   - Windows 2000
'   - Windows ME
'¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯
'  FindFirstUrlCacheGroup
'  FindNextUrlCacheGroup
'  FtpCommand
'  FtpGetFileSize
'  GetUrlCacheGroupAttribute
'  InternetGetConnectedStateEx
'  SetUrlCacheGroupAttribute
'
'_____________________________________________________________________________________________________________
' The following functions are either not implemented, or are not supported:
'¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯
'  InternetInitializeAutoProxyDll
'  InternetSetDialState
'  InternetSetOptionEx
'_____________________________________________________________________________________________________________
'¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯
'-------------------------------------------------------------------------------------------------------------
' The following are type definitions that are necisary to understand because when making the transfer from
' C (Win32 API) to Visual Basic, you need to know the data type's size in bytes to match it up with the
' correct corisponding VB data type.
'-------------------------------------------------------------------------------------------------------------
' typedef __int64 LONGLONG;   // __int64 is a 64-bit (8-byte) integer
' typedef LONGLONG GROUPID;   // GROUPID = LONGLONG = __int64 which is a 64-bit (8-byte) integer
' typedef WORD INTERNET_PORT; // WORD is a 16-bit (2-byte) integer
'-------------------------------------------------------------------------------------------------------------
 



' Type Declarations
Public Type FILETIME ' *NOTE: Because the FILETIME structure is simply two 4 byte values stuck together, it can (and should) be replaced by the VB data type "Currency" which is it's equal in size (8 bytes)
  dwLowDateTime  As Long 'DWORD // Specifies the low-order 32 bits of the file time.
  dwHighDateTime As Long 'DWORD // Specifies the high-order 32 bits of the file time.
End Type

Public Type INTERNET_CACHE_ENTRY_INFO
  dwStructSize      As Long     'DWORD    // Unsigned long integer value that contains the size, in TCHARs, of this structure. This value can be used to help determine the version of the cache system.
  lpszSourceUrlName As Long     'LPTSTR   // Address of a string that contains the URL name. The string occupies the memory area at the end of this structure.
  lpszLocalFileName As Long     'LPTSTR   // Address of a string that contains the local file name. The string occupies the memory area at the end of this structure.
  CacheEntryType    As Long     'DWORD    // Unsigned long integer value that contains the cache type bitmask. Currently, the cache entry type value of resources from the Internet is equal to zero.
  'For History and Cookie entries, the cache entry type is a combination of two values.
  'One value determines how the cache entry is handled;
  'the second value indicates what is being cached.
  'The value that determines how the cache entry is handled can be one of the following:
'                                            EDITED_CACHE_ENTRY        = Cache entry has been altered since it was downloaded from the Internet.
'                                            NORMAL_CACHE_ENTRY        = Normal cache entry; can be deleted to recover space for new entries.
'                                            SPARSE_CACHE_ENTRY        = Not currently implemented.
'                                            STICKY_CACHE_ENTRY        = Sticky cache entry that is exempt from scavenging for the amount of time specified by dwExemptDelta. The default value set by CommitUrlCacheEntry is one day.
'                                            TRACK_OFFLINE_CACHE_ENTRY = The value that indicates what is being cached can be one of the following:
'                                            TRACK_ONLINE_CACHE_ENTRY  = The value that indicates what is being cached can be one of the following:
'                                            COOKIE_CACHE_ENTRY        = Cookie cache entry.
'                                            URLHISTORY_CACHE_ENTRY    = Visited link cache entry.
  dwUseCount        As Long     'DWORD    // Unsigned long integer value that contains the current user count of the cache entry.
  dwHitRate         As Long     'DWORD    // Unsigned long integer value that contains the number of times the cache entry was retrieved.
  dwSizeLow         As Long     'DWORD    // Unsigned long integer value that contains the low order of the file size, in TCHARs.
  dwSizeHigh        As Long     'DWORD    // Unsigned long integer value that contains the high-order DWORD of the file size, in TCHARs.
  LastModifiedTime  As Currency 'FILETIME // FILETIME structure that contains the last modified time of this URL, in Greenwich mean time format.
  ExpireTime        As Currency 'FILETIME // FILETIME structure that contains the expiration time of this file, in Greenwich mean time format.
  LastAccessTime    As Currency 'FILETIME // FILETIME structure that contains the last accessed time, in Greenwich mean time format.
  LastSyncTime      As Currency 'FILETIME // FILETIME structure that contains the last time the cache was synchronized.
  lpHeaderInfo      As Long     'LPBYTE   // Address of a buffer that contains the header information. The buffer occupies the memory at the end of this structure.
  dwHeaderInfoSize  As Long     'DWORD    // Unsigned long integer value that contains the size of the lpHeaderInfo buffer, in TCHARs.
  lpszFileExtension As Long     'LPTSTR   // Address of a string that contains the file extension used to retrieve the data as a file. The string occupies the memory area at the end of this structure.
  dwReserved        As Long     'DWORD    // Reserved. Must be set to zero.
  dwExemptDelta     As Long     'DWORD    // Unsigned long integer value that contains the exemption time, in seconds, from the last accessed time.
End Type

Public Type WIN32_FIND_DATA
  dwFileAttributes   As Long         'DWORD    // Specifies the file attributes of the file found. This member can be one or more of the constant values starting with FILE_ATTRIBUTE_*
  ftCreationTime     As Currency     'FILETIME // Specifies a FILETIME structure containing the time the file was created. FindFirstFile and FindNextFile report file times in Coordinated Universal Time (UTC) format. These functions set the FILETIME members to zero if the file system containing the file does not support this time member. You can use the FileTimeToLocalFileTime function to convert from UTC to local time, and then use the FileTimeToSystemTime function to convert the local time to a SYSTEMTIME structure containing individual members for the month, day, year, weekday, hour, minute, second, and millisecond.
  ftLastAccessTime   As Currency     'FILETIME // Specifies a FILETIME structure containing the time that the file was last accessed. The time is in UTC format; the FILETIME members are zero if the file system does not support this time member.
  ftLastWriteTime    As Currency     'FILETIME // Specifies a FILETIME structure containing the time that the file was last written to. The time is in UTC format; the FILETIME members are zero if the file system does not support this time member.
  nFileSizeHigh      As Long         'DWORD    // Specifies the high-order DWORD value of the file size, in bytes. This value is zero unless the file size is greater than MAXDWORD. The size of the file is equal to (nFileSizeHigh * (MAXDWORD+1)) + nFileSizeLow.
  nFileSizeLow       As Long         'DWORD    // Specifies the low-order DWORD value of the file size, in bytes.
  dwReserved0        As Long         'DWORD    // If the dwFileAttributes member includes the FILE_ATTRIBUTE_REPARSE_POINT attribute, this member specifies the reparse tag. Otherwise, this value is undefined and should not be used.
  dwReserved1        As Long         'DWORD    // Reserved for future use.
  cFileName          As String * 259 'TCHAR    // A null-terminated string that is the name of the file.
  cAlternateFileName As String * 13  'TCHAR    // A null-terminated string that is an alternative name for the file. This name is in the classic 8.3 (filename.ext) file name format.
End Type

Public Type GOPHER_FIND_DATA
  DisplayString        As String * 128 'TCHAR[MAX_GOPHER_DISPLAY_TEXT  + 1]   // Array of characters that contains the friendly name of an object. An application can display this string to allow the user to select the object.
  GopherType           As Long         'DWORD                                 // Unsigned long integer value that contains the mask of flags that describe the item returned. This can be one of the following values that starts with GOPHER_TYPE_*
  SizeLow              As Long         'DWORD                                 // Unsigned long integer value that contains the low 32 bits of the file size.
  SizeHigh             As Long         'DWORD                                 // Unsigned long integer value that contains the high 32 bits of the file size.
  LastModificationTime As Currency     'FILETIME                              // FILETIME value that contains the time when the file was last modified.
  Locator              As String * 652 'TCHAR [MAX_GOPHER_LOCATOR_LENGTH + 1] // Array of characters that identifies the file. An application can pass the locator string to GopherOpenFile or GopherFindFirstFile.
End Type

Public Type INTERNET_BUFFERS
  dwStructSize    As Long   'DWORD              // Unsigned long integer value used for API versioning. This is set to the size of the INTERNET_BUFFERS structure.
  Next            As Long   'INTERNET_BUFFERS * // Address of the next INTERNET_BUFFERS structure. (Use the VB "VarPtr" function to get the address of the INTERNET_BUFFERS variable to pass to this member)
  lpcszHeader     As String 'LPCTSTR            // Address of a string value that contains the headers. This value can be NULL.
  dwHeadersLength As Long   'DWORD              // Unsigned long integer value that contains the length of the headers, in TCHARs, if lpcszHeader is not NULL.
  dwHeadersTotal  As Long   'DWORD              // Unsigned long integer value that contains the size of the headers if there is not enough memory in the buffer.
  lpvBuffer       As Long   'LPVOID             // Address of the data buffer.
  dwBufferLength  As Long   'DWORD              // Unsigned long integer value that contains the length of the buffer, in TCHARs, if lpvBuffer is not NULL.
  dwBufferTotal   As Long   'DWORD              // Unsigned long integer value that contains the total size of the resource.
  dwOffsetLow     As Long   'DWORD              // Unsigned long integer value that is used for read ranges.
  dwOffsetHigh    As Long   'DWORD              // Unsigned long integer value that is used for read ranges.
End Type

' Enumeration - Possible values for the URL_COMPONENTS.nScheme member
Public Enum INTERNET_SCHEME
  INTERNET_SCHEME_PARTIAL = -2
  INTERNET_SCHEME_UNKNOWN = -1
  INTERNET_SCHEME_DEFAULT = 0
  INTERNET_SCHEME_FTP = 1
  INTERNET_SCHEME_GOPHER = 2
  INTERNET_SCHEME_HTTP = 3
  INTERNET_SCHEME_HTTPS = 4
  INTERNET_SCHEME_FILE = 5
  INTERNET_SCHEME_NEWS = 6
  INTERNET_SCHEME_MAILTO = 7
  INTERNET_SCHEME_SOCKS = 8
  INTERNET_SCHEME_JAVASCRIPT = 9
  INTERNET_SCHEME_VBSCRIPT = 10
  INTERNET_SCHEME_FIRST = INTERNET_SCHEME_FTP
  INTERNET_SCHEME_LAST = INTERNET_SCHEME_VBSCRIPT
End Enum

Public Type URL_COMPONENTS
  dwStructSize      As Long            'DWORD           // Size of this structure. Used in version check
  lpszScheme        As String          'LPSTR           // Pointer to scheme name
  dwSchemeLength    As Long            'DWORD           // Length of scheme name
  nScheme           As INTERNET_SCHEME 'INTERNET_SCHEME // Enumerated scheme type (if known)
  lpszHostName      As String          'LPSTR           // Pointer to host name
  dwHostNameLength  As Long            'DWORD           // Length of host name
  nPort             As Integer         'INTERNET_PORT   // Converted port number
  lpszUserName      As String          'LPSTR           // Pointer to user name
  dwUserNameLength  As Long            'DWORD           // Length of user name
  lpszPassword      As String          'LPSTR           // Pointer to password
  dwPasswordLength  As Long            'DWORD           // Length of password
  lpszUrlPath       As String          'LPSTR           // Pointer to URL-path
  dwUrlPathLength   As Long            'DWORD           // Length of URL-path
  lpszExtraInfo     As String          'LPSTR           // Pointer to extra information (e.g. ?foo or #foo)
  dwExtraInfoLength As Long            'DWORD           // Length of extra information
End Type

Public Type INTERNET_ASYNC_RESULT
  dwResult As Long 'DWORD // Unsigned long integer value that references an HINTERNET handle, unsigned long integer, or Boolean return code from an asynchronous function.
                   '         *NOTE: If dwInternetStatus = INTERNET_STATUS_HANDLE_CREATED or INTERNET_STATUS_REQUEST_COMPLETE then dwResult (this member) is the address of the HINTERNET handle
  dwError  As Long 'DWORD // Unsigned long integer value that contains the error message if dwResult indicates that the function failed. If the operation succeeded, this member usually contains ERROR_SUCCESS.
End Type

Public Type sockaddr
  sa_Family As Integer     'u_short  // Address family
  sa_Data   As String * 13 'char[14] // Up to 14 bytes of direct address
End Type

Public Type REQUEST_CONTEXT
  hWindow    As Long         'HWND      // Main window handle
  nURL       As Long         'int       // ID of the edit box with the URL
  nHeader    As Long         'int       // ID of the edit box for the header info
  nResource  As Long         'int       // ID of the edit box for the resource
  hOpen      As Long         'HINTERNET // HINTERNET handle created by InternetOpen
  hResource  As Long         'HINTERNET // HINTERNET handle created by InternetOpenUrl
  szMemo     As String * 511 'char[512] // String to store status memo
  hThread    As Long         'HANDLE    // Thread handle
  dwThreadID As Long         'DWORD     // Thread ID
End Type

Public Type SYSTEMTIME
  wYear         As Integer 'WORD // Specifies the current year.
  wMonth        As Integer 'WORD // Specifies the current month; January = 1, February = 2, and so on.
  wDayOfWeek    As Integer 'WORD // Specifies the current day of the week; Sunday = 0, Monday = 1, and so on.
  wDay          As Integer 'WORD // Specifies the current day of the month.
  wHour         As Integer 'WORD // Specifies the current hour.
  wMinute       As Integer 'WORD // Specifies the current minute.
  wSecond       As Integer 'WORD // Specifies the current second.
  wMilliseconds As Integer 'WORD // Specifies the current millisecond.
End Type

Public Type GOPHER_ADMIN_ATTRIBUTE
  Comment      As String 'LPCTSTR
  EmailAddress As String 'LPCTSTR
End Type

Public Type GOPHER_MOD_DATE_ATTRIBUTE
  DateAndTime As Currency 'FILETIME
End Type

Public Type GOPHER_TTL_ATTRIBUTE
  Ttl As Long 'DWORD
End Type

Public Type GOPHER_SCORE_ATTRIBUTE
  Score As Long 'INT
End Type

Public Type GOPHER_SCORE_RANGE_ATTRIBUTE
  LowerBound As Long 'INT
  UpperBound As Long 'INT
End Type

Public Type GOPHER_SITE_ATTRIBUTE
  Site As String 'LPCTSTR
End Type

Public Type GOPHER_ORGANIZATION_ATTRIBUTE
  Organization As String 'LPCTSTR
End Type

Public Type GOPHER_LOCATION_ATTRIBUTE
  Location As String 'LPCTSTR
End Type

Public Type GOPHER_GEOGRAPHICAL_LOCATION_ATTRIBUTE
  DegreesNorth As Long 'INT
  MinutesNorth As Long 'INT
  SecondsNorth As Long 'INT
  DegreesEast  As Long 'INT
  MinutesEast  As Long 'INT
  SecondsEast  As Long 'INT
End Type

Public Type GOPHER_TIMEZONE_ATTRIBUTE
  Zone As Long 'INT
End Type

Public Type GOPHER_PROVIDER_ATTRIBUTE
  Provider As String 'LPCTSTR
End Type

Public Type GOPHER_VERSION_ATTRIBUTE
  Version As String 'LPCTSTR
End Type

Public Type GOPHER_ABSTRACT_ATTRIBUTE
  ShortAbstract As String 'LPCTSTR
  AbstractFile As String  'LPCTSTR
End Type

Public Type GOPHER_VIEW_ATTRIBUTE
  ContentType As String 'LPCTSTR
  Language    As String 'LPCTSTR
  Size        As Long   'DWORD
End Type

Public Type GOPHER_VERONICA_ATTRIBUTE
  TreeWalk As Long 'BOOL
End Type

Public Type GOPHER_ASK_ATTRIBUTE
  QuestionType As String 'LPCTSTR
  QuestionText As String 'LPCTSTR
End Type

Public Type GOPHER_UNKNOWN_ATTRIBUTE ' This is returned if we retrieve an attribute that is not specified in the current gopher/gopher+ documentation. It is up to the application to parse the information
  Text As String 'LPCTSTR
End Type

Public Type GOPHER_ATTRIBUTE
  CategoryId           As Long 'DWORD
  AttributeId          As Long 'DWORD
  Admin                As GOPHER_ADMIN_ATTRIBUTE
  ModDate              As GOPHER_MOD_DATE_ATTRIBUTE
  Score                As GOPHER_SCORE_ATTRIBUTE
  ScoreRange           As GOPHER_SCORE_RANGE_ATTRIBUTE
  Site                 As GOPHER_SITE_ATTRIBUTE
  Organization         As GOPHER_ORGANIZATION_ATTRIBUTE
  Location             As GOPHER_LOCATION_ATTRIBUTE
  GeographicalLocation As GOPHER_GEOGRAPHICAL_LOCATION_ATTRIBUTE
  TimeZone             As GOPHER_TIMEZONE_ATTRIBUTE
  Provider             As GOPHER_PROVIDER_ATTRIBUTE
  Version              As GOPHER_VERSION_ATTRIBUTE
  Abstract             As GOPHER_ABSTRACT_ATTRIBUTE
  View                 As GOPHER_VIEW_ATTRIBUTE
  Veronica             As GOPHER_VERONICA_ATTRIBUTE
  Ask                  As GOPHER_ASK_ATTRIBUTE
  Unknown              As GOPHER_UNKNOWN_ATTRIBUTE
End Type

' Constants - WIN32_FIND_DATA.dwFileAttributes
Private Enum FileAttributes
  FILE_ATTRIBUTE_ARCHIVE = &H20        ' The file or directory is an archive file or directory. Applications use this attribute to mark files for backup or removal.
  FILE_ATTRIBUTE_COMPRESSED = &H800    ' The file or directory is compressed. For a file, this means that all of the data in the file is compressed. For a directory, this means that compression is the default for newly created files and subdirectories.
  FILE_ATTRIBUTE_DIRECTORY = &H10      ' The handle identifies a directory.
  FILE_ATTRIBUTE_ENCRYPTED = &H4000    ' The file or directory is encrypted. For a file, this means that all data in the file is encrypted. For a directory, this means that encryption is the default for newly created files and subdirectories.
  FILE_ATTRIBUTE_HIDDEN = &H2          ' The file or directory is hidden. It is not included in an ordinary directory listing.
  FILE_ATTRIBUTE_NORMAL = &H80         ' The file or directory has no other attributes set. This attribute is valid only if used alone.
  FILE_ATTRIBUTE_OFFLINE = &H1000      ' The file data is not immediately available. This attribute indicates that the file data has been physically moved to offline storage. This attribute is used by Remote Storage, the hierarchical storage management software in Windows 2000. Applications should not arbitrarily change this attribute.
  FILE_ATTRIBUTE_READONLY = &H1        ' The file or directory is read-only. Applications can read the file but cannot write to it or delete it. In the case of a directory, applications cannot delete it.
  FILE_ATTRIBUTE_SYSTEM = &H4          ' The file or directory is part of the operating system or is used exclusively by the operating system.
  FILE_ATTRIBUTE_TEMPORARY = &H100     ' The file is being used for temporary storage. File systems attempt to keep all of the data in memory for quicker access, rather than flushing it back to mass storage. A temporary file should be deleted by the application as soon as it is no longer needed.
End Enum

' Constants - General
Public Const MAX_PATH = 260
Public Const MAXCHAR = &H7F
Public Const MAXSHORT = &H7FFF
Public Const MAXLONG = &H7FFFFFFF
Public Const MAXBYTE = &HFF
Public Const MAXWORD = &HFFFF
Public Const MAXDWORD = &HFFFFFFFF

' Constants - Cache entry type flags:
Public Const NORMAL_CACHE_ENTRY        As Long = &H1      ' Normal cache entry; can be deleted to recover space for new entries.
Public Const STICKY_CACHE_ENTRY        As Long = &H4      ' Sticky cache entry; exempt from scavenging.
Public Const EDITED_CACHE_ENTRY        As Long = &H8      '
Public Const TRACK_OFFLINE_CACHE_ENTRY As Long = &H10     ' Not currently implemented.
Public Const TRACK_ONLINE_CACHE_ENTRY  As Long = &H20     ' Not currently implemented.
Public Const SPARSE_CACHE_ENTRY        As Long = &H10000  '
Public Const COOKIE_CACHE_ENTRY        As Long = &H100000 ' Cookie cache entry.
Public Const URLHISTORY_CACHE_ENTRY    As Long = &H200000 ' Visited link cache entry.

' Constants - FindFirstUrlCacheGroup.dwFilter
Public Const CACHEGROUP_SEARCH_ALL   As Long = &H0 ' Search all of the cache groups.
Public Const CACHEGROUP_SEARCH_BYURL As Long = &H1 ' Not currently implemented.

' Constants - FtpCommand.dwFlags
'Public Const FTP_TRANSFER_TYPE_ASCII = &H1  ' Transfers the file using FTP's ASCII (Type A) transfer method. Control and formatting information is converted to local equivalents.
'Public Const FTP_TRANSFER_TYPE_BINARY = &H2 ' Transfers the file using FTP's Image (Type I) transfer method. The file is transferred exactly as it exists with no changes. This is the default transfer method.

' Constants - FtpFindFirstFile.dwFlags
'Public Const INTERNET_FLAG_HYPERLINK = &H400          ' Asking wininet to do hyperlinking semantic which works right for scripts
'Public Const INTERNET_FLAG_NEED_FILE = &H10           ' Need a file for this request
'Public Const INTERNET_FLAG_NO_CACHE_WRITE = &H4000000 ' Don't write this item to the cache
'Public Const INTERNET_FLAG_RELOAD = &H80000000        ' Retrieve the original item
'Public Const INTERNET_FLAG_RESYNCHRONIZE = &H800      ' Asking wininet to update an item if it is newer

' Constants - FtpGetFile.dwFlags / FtpOpenFile.dwFlags / FtpPutFile.dwFlags / GopherFindFirstFile.dwFlags / GopherOpenFile.dwFlags
Private Enum FileTransferTypes
  FTP_TRANSFER_TYPE_ASCII = &H1       ' Transfers the file using FTP's ASCII (Type A) transfer method. Control and formatting information is converted to local equivalents.
  FTP_TRANSFER_TYPE_BINARY = &H2      ' Transfers the file using FTP's Image (Type I) transfer method. The file is transferred exactly as it exists with no changes. This is the default transfer method.
  FTP_TRANSFER_TYPE_UNKNOWN = &H0     ' Defaults to FTP_TRANSFER_TYPE_BINARY.
  INTERNET_FLAG_TRANSFER_ASCII = &H1  ' Transfers the file as ASCII.
  INTERNET_FLAG_TRANSFER_BINARY = &H2 ' Transfers the file as binary.
  INTERNET_FLAG_HYPERLINK = &H400     ' Forces a reload if there was no Expires time and no LastModified time returned from the server when determining whether to reload the item from the network.
  INTERNET_FLAG_NEED_FILE = &H10      ' Causes a temporary file to be created if the file cannot be cached.
  INTERNET_FLAG_RELOAD = &H80000000   ' Forces a download of the requested file, object, or directory listing from the origin server, not from the cache.
  INTERNET_FLAG_RESYNCHRONIZE = &H800 ' Reloads HTTP resources if the resource has been modified since the last time it was downloaded. All FTP and Gopher resources are reloaded.
End Enum

' Constants - FtpOpenFile.dwAccess
Public Const GENERIC_READ = &H80000000
Public Const GENERIC_WRITE = &H40000000

' Constants - GetUrlCacheGroupAttribute.dwAttributes
Public Const CACHEGROUP_ATTRIBUTE_BASIC = &H1          ' Retrieves the flags, type, and disk quota attributes of the cache group.
Public Const CACHEGROUP_ATTRIBUTE_FLAG = &H2           ' Sets or retrieves the flags associated with the cache group.
Public Const CACHEGROUP_ATTRIBUTE_GET_ALL = &HFFFFFFFF ' Retrieves all the attributes of the cache group.
Public Const CACHEGROUP_ATTRIBUTE_GROUPNAME = &H10     ' Sets or retrieves the group name of the cache group.
Public Const CACHEGROUP_ATTRIBUTE_QUOTA = &H8          ' Sets or retrieves the disk quota associated with the cache group.
Public Const CACHEGROUP_ATTRIBUTE_STORAGE = &H20       ' Sets or retrieves the group owner storage associated with the cache group.
Public Const CACHEGROUP_ATTRIBUTE_TYPE = &H4           ' Sets or retrieves the cache group type.

' Constants - SetUrlCacheGroupAttribute.dwAttributes
'Public Const CACHEGROUP_ATTRIBUTE_FLAG = &H2       ' Sets or retrieves the flags associated with the cache group.
'Public Const CACHEGROUP_ATTRIBUTE_GROUPNAME = &H10 ' Sets or retrieves the group name of the cache group.
'Public Const CACHEGROUP_ATTRIBUTE_QUOTA = &H8      ' Sets or retrieves the disk quota associated with the cache group.
'Public Const CACHEGROUP_ATTRIBUTE_STORAGE = &H20   ' Sets or retrieves the group owner storage associated with the cache group.
'Public Const CACHEGROUP_ATTRIBUTE_TYPE = &H4       ' Sets or retrieves the cache group type.
Public Const CACHEGROUP_READWRITE_MASK = CACHEGROUP_ATTRIBUTE_TYPE Or _
                                         CACHEGROUP_ATTRIBUTE_QUOTA Or _
                                         CACHEGROUP_ATTRIBUTE_GROUPNAME Or _
                                         CACHEGROUP_ATTRIBUTE_STORAGE ' Sets the type, disk quota, group name, and owner storage attributes of the cache group.

' Constants - GOPHER_FIND_DATA.GopherType
Public Const GOPHER_TYPE_ASK = &H40000000         ' Ask+ item.
Public Const GOPHER_TYPE_BINARY = &H200           ' Binary file.
Public Const GOPHER_TYPE_BITMAP = &H4000          ' Bitmap file.
Public Const GOPHER_TYPE_CALENDAR = &H80000       ' Calendar file.
Public Const GOPHER_TYPE_CSO = &H4                ' CSO telephone book server.
Public Const GOPHER_TYPE_DIRECTORY = &H2          ' Directory of additional Gopher items.
Public Const GOPHER_TYPE_DOS_ARCHIVE = &H20       ' MS-DOS archive file.
Public Const GOPHER_TYPE_ERROR = &H8              ' Indicator of an error condition.
Public Const GOPHER_TYPE_GIF = &H1000             ' GIF graphics file.
Public Const GOPHER_TYPE_GOPHER_PLUS = &H80000000 ' Gopher+ item.
Public Const GOPHER_TYPE_HTML = &H20000           ' HTML document.
Public Const GOPHER_TYPE_IMAGE = &H2000           ' Image file.
Public Const GOPHER_TYPE_INDEX_SERVER = &H80      ' Index server.
Public Const GOPHER_TYPE_INLINE = &H100000        ' Inline file.
Public Const GOPHER_TYPE_MAC_BINHEX = &H10        ' Macintosh file in BINHEX format.
Public Const GOPHER_TYPE_MOVIE = &H8000           ' Movie file.
Public Const GOPHER_TYPE_PDF = &H40000            ' PDF file.
Public Const GOPHER_TYPE_REDUNDANT = &H400        ' Indicator of a duplicated server. The information contained within is a duplicate of the primary server. The primary server is defined as the last directory entry that did not have a GOPHER_TYPE_REDUNDANT type.
Public Const GOPHER_TYPE_SOUND = &H10000          ' Sound file.
Public Const GOPHER_TYPE_TELNET = &H100           ' Telnet server.
Public Const GOPHER_TYPE_TEXT_FILE = &H1          ' ASCII text file.
Public Const GOPHER_TYPE_TN3270 = &H800           ' TN3270 server.
Public Const GOPHER_TYPE_UNIX_UUENCODED = &H40    ' UUENCODED file.
Public Const GOPHER_TYPE_UNKNOWN = &H20000000     ' Item type is unknown.

' Constants - GopherGetAttribute.lpfnEnumerator(GOPHER_ATTRIBUTE_TYPE).CategoryId
Public Const GOPHER_ATTRIBUTE_ID_BASE = &HABCCCC00
Public Const GOPHER_CATEGORY_ID_ALL = (GOPHER_ATTRIBUTE_ID_BASE + 1)
Public Const GOPHER_CATEGORY_ID_INFO = (GOPHER_ATTRIBUTE_ID_BASE + 2)
Public Const GOPHER_CATEGORY_ID_ADMIN = (GOPHER_ATTRIBUTE_ID_BASE + 3)
Public Const GOPHER_CATEGORY_ID_VIEWS = (GOPHER_ATTRIBUTE_ID_BASE + 4)
Public Const GOPHER_CATEGORY_ID_ABSTRACT = (GOPHER_ATTRIBUTE_ID_BASE + 5)
Public Const GOPHER_CATEGORY_ID_VERONICA = (GOPHER_ATTRIBUTE_ID_BASE + 6)
Public Const GOPHER_CATEGORY_ID_ASK = (GOPHER_ATTRIBUTE_ID_BASE + 7)
Public Const GOPHER_CATEGORY_ID_UNKNOWN = (GOPHER_ATTRIBUTE_ID_BASE + 8)

' Constants - GopherGetAttribute.lpfnEnumerator(GOPHER_ATTRIBUTE_TYPE).AttributeId
Public Const GOPHER_ATTRIBUTE_ID_ALL = (GOPHER_ATTRIBUTE_ID_BASE + 9)
Public Const GOPHER_ATTRIBUTE_ID_ADMIN = (GOPHER_ATTRIBUTE_ID_BASE + 10)
Public Const GOPHER_ATTRIBUTE_ID_MOD_DATE = (GOPHER_ATTRIBUTE_ID_BASE + 11)
Public Const GOPHER_ATTRIBUTE_ID_TTL = (GOPHER_ATTRIBUTE_ID_BASE + 12)
Public Const GOPHER_ATTRIBUTE_ID_SCORE = (GOPHER_ATTRIBUTE_ID_BASE + 13)
Public Const GOPHER_ATTRIBUTE_ID_RANGE = (GOPHER_ATTRIBUTE_ID_BASE + 14)
Public Const GOPHER_ATTRIBUTE_ID_SITE = (GOPHER_ATTRIBUTE_ID_BASE + 15)
Public Const GOPHER_ATTRIBUTE_ID_ORG = (GOPHER_ATTRIBUTE_ID_BASE + 16)
Public Const GOPHER_ATTRIBUTE_ID_LOCATION = (GOPHER_ATTRIBUTE_ID_BASE + 17)
Public Const GOPHER_ATTRIBUTE_ID_GEOG = (GOPHER_ATTRIBUTE_ID_BASE + 18)
Public Const GOPHER_ATTRIBUTE_ID_TIMEZONE = (GOPHER_ATTRIBUTE_ID_BASE + 19)
Public Const GOPHER_ATTRIBUTE_ID_PROVIDER = (GOPHER_ATTRIBUTE_ID_BASE + 20)
Public Const GOPHER_ATTRIBUTE_ID_VERSION = (GOPHER_ATTRIBUTE_ID_BASE + 21)
Public Const GOPHER_ATTRIBUTE_ID_ABSTRACT = (GOPHER_ATTRIBUTE_ID_BASE + 22)
Public Const GOPHER_ATTRIBUTE_ID_VIEW = (GOPHER_ATTRIBUTE_ID_BASE + 23)
Public Const GOPHER_ATTRIBUTE_ID_TREEWALK = (GOPHER_ATTRIBUTE_ID_BASE + 24)
Public Const GOPHER_ATTRIBUTE_ID_UNKNOWN = (GOPHER_ATTRIBUTE_ID_BASE + 25)

' Constants - HttpAddRequestHeaders.dwModifiers
Public Const HTTP_ADDREQ_FLAG_ADD = &H20000000                    ' Adds the header if it does not exist. Used with HTTP_ADDREQ_FLAG_REPLACE.
Public Const HTTP_ADDREQ_FLAG_ADD_IF_NEW = &H10000000             ' Adds the header only if it does not already exist; otherwise, an error is returned.
Public Const HTTP_ADDREQ_FLAG_COALESCE = &H40000000               ' Coalesces headers of the same name.
Public Const HTTP_ADDREQ_FLAG_COALESCE_WITH_COMMA = &H40000000    ' Coalesces headers of the same name. For example, adding "Accept: text/*" followed by "Accept: audio/*" with this flag results in the formation of the single header "Accept: text/*, audio/*". This causes the first header found to be coalesced. It is up to the calling application to ensure a cohesive scheme with respect to coalesced/separate headers.
Public Const HTTP_ADDREQ_FLAG_COALESCE_WITH_SEMICOLON = &H1000000 ' Coalesces headers of the same name using a semicolon.
Public Const HTTP_ADDREQ_FLAG_REPLACE = &H80000000                ' Replaces or removes a header. If the header value is empty and the header is found, it is removed. If not empty, the header value is replaced.

' Constants - HttpEndRequest.dwFlags / HttpSendRequestEx.dwFlags
Public Const HSR_ASYNC = &H1       ' Forces asynchronous operations.
Public Const HSR_SYNC = &H4        ' Forces synchronous operations.
Public Const HSR_USE_CONTEXT = &H8 ' Forces HttpEndRequest to use the context value, even if it is set to zero.
Public Const HSR_INITIATE = &H8    ' Iterative operation (completed by HttpEndRequest).
Public Const HSR_DOWNLOAD = &H10   ' Download resource to file.
Public Const HSR_CHUNKED = &H20    ' Send chunked data.

' Constants - HttpOpenRequest.dwFlags
Public Const INTERNET_FLAG_CACHE_IF_NET_FAIL = &H10000        ' Returns the resource from the cache if the network request for the resource fails due to an ERROR_INTERNET_CONNECTION_RESET (the connection with the server has been reset) or ERROR_INTERNET_CANNOT_CONNECT (the attempt to connect to the server failed).
'Public Const INTERNET_FLAG_HYPERLINK = &H400                 ' Forces a reload if there was no Expires time and no LastModified time returned from the server when determining whether to reload the item from the network.
'Public Const INTERNET_FLAG_IGNORE_CERT_CN_INVALID = &H1000   ' Disables Microsoft® Win32® Internet function checking of SSL/PCT-based certificates that are returned from the server against the host name given in the request. Win32 Internet functions use a simple check against certificates by comparing for matching host names and simple wildcarding rules.
'Public Const INTERNET_FLAG_IGNORE_CERT_DATE_INVALID = &H2000 ' Disables Win32 Internet function checking of SSL/PCT-based certificates for proper validity dates.
Public Const INTERNET_FLAG_IGNORE_REDIRECT_TO_HTTP = &H8000   ' Disables the ability of the Win32 Internet functions to detect this special type of redirect. When this flag is used, Win32 Internet functions transparently allow redirects from HTTPS to HTTP URLs.
Public Const INTERNET_FLAG_IGNORE_REDIRECT_TO_HTTPS = &H4000  ' Disables the ability of the Win32 Internet functions to detect this special type of redirect. When this flag is used, Win32 Internet functions transparently allow redirects from HTTP to HTTPS URLs.
Public Const INTERNET_FLAG_KEEP_CONNECTION = &H400000         ' Uses keep-alive semantics, if available, for the connection. This flag is required for Microsoft Network (MSN), NT LAN Manager (NTLM), and other types of authentication.
'Public Const INTERNET_FLAG_NEED_FILE = &H10                  ' Causes a temporary file to be created if the file cannot be cached.
Public Const INTERNET_FLAG_NO_AUTH = &H40000                  ' Does not attempt authentication automatically.
Public Const INTERNET_FLAG_NO_AUTO_REDIRECT = &H200000        ' Does not automatically handle redirection in HttpSendRequest.
Public Const INTERNET_FLAG_NO_CACHE_WRITE = &H4000000         ' Does not add the returned entity to the cache.
Public Const INTERNET_FLAG_NO_COOKIES = &H80000               ' Does not automatically add cookie headers to requests, and does not automatically add returned cookies to the cookie database.
Public Const INTERNET_FLAG_NO_UI = &H200                      ' Disables the cookie dialog box.
Public Const INTERNET_FLAG_PRAGMA_NOCACHE = &H100             ' Forces the request to be resolved by the origin server, even if a cached copy exists on the proxy.
'Public Const INTERNET_FLAG_RELOAD = &H80000000               ' Forces a download of the requested file, object, or directory listing from the origin server, not from the cache.
'Public Const INTERNET_FLAG_RESYNCHRONIZE = &H800             ' Reloads HTTP resources if the resource has been modified since the last time it was downloaded. All FTP and Gopher resources are reloaded.
Public Const INTERNET_FLAG_SECURE = &H800000                  ' Uses secure transaction semantics. This translates to using Secure Sockets Layer/Private Communications Technology (SSL/PCT) and is only meaningful in HTTP requests.

' Constants - HttpQueryInfo.dwInfoLevel
Public Const HTTP_QUERY_ACCEPT = 24                   ' Retrieves the acceptable media types for the response.
Public Const HTTP_QUERY_ACCEPT_CHARSET = 25           ' Retrieves the acceptable character sets for the response.
Public Const HTTP_QUERY_ACCEPT_ENCODING = 26          ' Retrieves the acceptable content-coding values for the response.
Public Const HTTP_QUERY_ACCEPT_LANGUAGE = 27          ' Retrieves the acceptable natural languages for the response.
Public Const HTTP_QUERY_ACCEPT_RANGES = 42            ' Retrieves the types of range requests that are accepted for a resource.
Public Const HTTP_QUERY_AGE = 48                      ' Retrieves the Age response-header field, which contains the sender's estimate of the amount of time since the response was generated at the origin server.
Public Const HTTP_QUERY_ALLOW = 7                     ' Receives the methods supported by the server.
Public Const HTTP_QUERY_AUTHORIZATION = 28            ' Retrieves the authorization credentials used for a request.
Public Const HTTP_QUERY_CACHE_CONTROL = 49            ' Retrieves the cache control directives.
Public Const HTTP_QUERY_CONNECTION = 23               ' Retrieves any options that are specified for a particular connection and must not be communicated by proxies over further connections.
Public Const HTTP_QUERY_CONTENT_BASE = 50             ' Retrieves the base URI = Uniform Resource Identifier) for resolving relative URLs within the entity.
Public Const HTTP_QUERY_CONTENT_DESCRIPTION = 4       ' Obsolete. Maintained for legacy application compatibility only.
Public Const HTTP_QUERY_CONTENT_DISPOSITION = 47      ' Obsolete. Maintained for legacy application compatibility only.
Public Const HTTP_QUERY_CONTENT_ENCODING = 29         ' Retrieves any additional content codings that have been applied to the entire resource.
Public Const HTTP_QUERY_CONTENT_ID = 3                ' Retrieves the content identification.
Public Const HTTP_QUERY_CONTENT_LANGUAGE = 6          ' Retrieves the language that the content is in.
Public Const HTTP_QUERY_CONTENT_LENGTH = 5            ' Retrieves the size of the resource, in bytes.
Public Const HTTP_QUERY_CONTENT_LOCATION = 51         ' Retrieves the resource location for the entity enclosed in the message.
Public Const HTTP_QUERY_CONTENT_MD5 = 52              ' Retrieves an MD5 digest of the entity-body for the purpose of providing an end-to-end message integrity check (MIC) for the entity-body. For more information, see RFC1864, The Content-MD5 Header Field, at ftp://ftp.isi.edu/in-notes/rfc1864.txt
Public Const HTTP_QUERY_CONTENT_RANGE = 53            ' Retrieves the location in the full entity-body where the partial entity-body should be inserted and the total size of the full entity-body.
Public Const HTTP_QUERY_CONTENT_TRANSFER_ENCODING = 2 ' Receives the additional content coding that has been applied to the resource.
Public Const HTTP_QUERY_CONTENT_TYPE = 1              ' Receives the content type of the resource (such as text/html).
Public Const HTTP_QUERY_COOKIE = 44                   ' Retrieves any cookies associated with the request.
Public Const HTTP_QUERY_COST = 15                     ' No longer supported.
Public Const HTTP_QUERY_CUSTOM = 65535                ' Causes HttpQueryInfo to search for the header name specified in lpvBuffer and store the header information in lpvBuffer.
Public Const HTTP_QUERY_DATE = 9                      ' Receives the date and time at which the message was originated.
Public Const HTTP_QUERY_DERIVED_FROM = 14             ' No longer supported.
Public Const HTTP_QUERY_ECHO_HEADERS = 73             ' Not currently implemented.
Public Const HTTP_QUERY_ECHO_HEADERS_CRLF = 74        ' Not currently implemented.
Public Const HTTP_QUERY_ECHO_REPLY = 72               ' Not currently implemented.
Public Const HTTP_QUERY_ECHO_REQUEST = 71             ' Not currently implemented.
Public Const HTTP_QUERY_ETAG = 54                     ' Retrieves the entity tag for the associated entity.
Public Const HTTP_QUERY_EXPECT = 68                   ' Retrieves the Expect header, which indicates whether the client application should expect 100 series responses.
Public Const HTTP_QUERY_EXPIRES = 10                  ' Receives the date and time after which the resource should be considered outdated.
Public Const HTTP_QUERY_FORWARDED = 30                ' Obsolete. Maintained for legacy application compatibility only.
Public Const HTTP_QUERY_FROM = 31                     ' Retrieves the e-mail address for the human user who controls the requesting user agent if the From header is given.
Public Const HTTP_QUERY_HOST = 55                     ' Retrieves the Internet host and port number of the resource being requested.
Public Const HTTP_QUERY_IF_MATCH = 56                 ' Retrieves the contents of the If-Match request-header field.
Public Const HTTP_QUERY_IF_MODIFIED_SINCE = 32        ' Retrieves the contents of the If-Modified-Since header.
Public Const HTTP_QUERY_IF_NONE_MATCH = 57            ' Retrieves the contents of the If-None-Match request-header field.
Public Const HTTP_QUERY_IF_RANGE = 58                 ' Retrieves the contents of the If-Range request-header field. This header allows the client application to check if the entity related to a partial copy of the entity in the client application's cache has not been updated. If the entity has not been updated, send the parts that the client application is missing. If the entity has been updated, send the entire updated entity.
Public Const HTTP_QUERY_IF_UNMODIFIED_SINCE = 59      ' Retrieves the contents of the If-Unmodified-Since request-header field.
Public Const HTTP_QUERY_LAST_MODIFIED = 11            ' Receives the date and time at which the server believes the resource was last modified.
Public Const HTTP_QUERY_LINK = 16                     ' Obsolete. Maintained for legacy application compatibility only.
Public Const HTTP_QUERY_LOCATION = 33                 ' Retrieves the absolute URI = Uniform Resource Identifier) used in a Location response-header.
Public Const HTTP_QUERY_MAX = 75                      ' Not a query flag. Indicates the maximum value of an HTTP_QUERY_* value.
Public Const HTTP_QUERY_MAX_FORWARDS = 60             ' Retrieves the number of proxies or gateways that can forward the request to the next inbound server.
Public Const HTTP_QUERY_MESSAGE_ID = 12               ' No longer supported.
Public Const HTTP_QUERY_MIME_VERSION = 0              ' Receives the version of the MIME protocol that was used to construct the message.
Public Const HTTP_QUERY_ORIG_URI = 34                 ' Obsolete. Maintained for legacy application compatibility only.
Public Const HTTP_QUERY_PRAGMA = 17                   ' Receives the implementation-specific directives that might apply to any recipient along the request/response chain.
Public Const HTTP_QUERY_PROXY_AUTHENTICATE = 41       ' Retrieves the authentication scheme and realm returned by the proxy.
Public Const HTTP_QUERY_PROXY_AUTHORIZATION = 61      ' Retrieves the header that is used to identify the user to a proxy that requires authentication. This header can only be retrieved before the request is sent to the server.
Public Const HTTP_QUERY_PROXY_CONNECTION = 69         ' Retrieves the Proxy-Connection header.
Public Const HTTP_QUERY_PUBLIC = 8                    ' Receives methods available at this server.
Public Const HTTP_QUERY_RANGE = 62                    ' Retrieves the byte range of an entity.
Public Const HTTP_QUERY_RAW_HEADERS = 21              ' Receives all the headers returned by the server. Each header is terminated by "\0". An additional "\0" terminates the list of headers.
Public Const HTTP_QUERY_RAW_HEADERS_CRLF = 22         ' Receives all the headers returned by the server. Each header is separated by a carriage return/line feed (CR/LF) sequence.
Public Const HTTP_QUERY_REFERER = 35                  ' Receives the URI (Uniform Resource Identifier) of the resource where the requested URI was obtained.
Public Const HTTP_QUERY_REFRESH = 46                  ' Obsolete. Maintained for legacy application compatibility only.
Public Const HTTP_QUERY_REQUEST_METHOD = 45           ' Receives the verb that is being used in the request, typically GET or POST.
Public Const HTTP_QUERY_RETRY_AFTER = 36              ' Retrieves the amount of time the service is expected to be unavailable.
Public Const HTTP_QUERY_SERVER = 37                   ' Retrieves information about the software used by the origin server to handle the request.
Public Const HTTP_QUERY_SET_COOKIE = 43               ' Receives the value of the cookie set for the request.
Public Const HTTP_QUERY_STATUS_CODE = 19              ' Receives the status code returned by the server. For a list of possible values, see HTTP Status Codes.
Public Const HTTP_QUERY_STATUS_TEXT = 20              ' Receives any additional text returned by the server on the response line.
Public Const HTTP_QUERY_TITLE = 38                    ' Obsolete. Maintained for legacy application compatibility only.
Public Const HTTP_QUERY_TRANSFER_ENCODING = 63        ' Retrieves the type of transformation that has been applied to the message body so it can be safely transferred between the sender and recipient.
Public Const HTTP_QUERY_UNLESS_MODIFIED_SINCE = 70    ' Retrieves the Unless-Modified-Since header.
Public Const HTTP_QUERY_UPGRADE = 64                  ' Retrieves the additional communication protocols that are supported by the server.
Public Const HTTP_QUERY_URI = 13                      ' Receives some or all of the Uniform Resource Identifiers (URIs) by which the Request-URI resource can be identified.
Public Const HTTP_QUERY_USER_AGENT = 39               ' Retrieves information about the user agent that made the request.
Public Const HTTP_QUERY_VARY = 65                     ' Retrieves the header that indicates that the entity was selected from a number of available representations of the response using server-driven negotiation.
Public Const HTTP_QUERY_VERSION = 18                  ' Receives the last response code returned by the server.
Public Const HTTP_QUERY_VIA = 66                      ' Retrieves the intermediate protocols and recipients between the user agent and the server on requests, and between the origin server and the client on responses.
Public Const HTTP_QUERY_WARNING = 67                  ' Retrieves additional information about the status of a response that might not be reflected by the response status code.
Public Const HTTP_QUERY_WWW_AUTHENTICATE = 40         ' Retrieves the authentication scheme and realm returned by the server.
Public Const HTTP_QUERY_FLAG_COALESCE = &H10000000    ' Not implemented.
Public Const HTTP_QUERY_FLAG_NUMBER = &H20000000      ' Returns the data as a 32-bit number for headers whose value is a number, such as the status code.
Public Const HTTP_QUERY_FLAG_REQUEST_HEADERS = &H80000000 ' Queries request headers only.
Public Const HTTP_QUERY_FLAG_SYSTEMTIME = &H40000000  ' Returns the header value as a standard Microsoft® Win32®SYSTEMTIME  structure, which does not require the application to parse the data. Use for headers whose value is a date/time string, such as "Last-Modified-Time".

' Constants - InternetAutodial.dwFlags / InternetDial.dwFlags
Public Const INTERNET_AUTODIAL_FORCE_ONLINE = 1        ' Forces an online Internet connection.
Public Const INTERNET_AUTODIAL_FORCE_UNATTENDED = 2    ' Forces an unattended Internet dial-up. If user intervention is required, the function will fail.
Public Const INTERNET_AUTODIAL_FAILIFSECURITYCHECK = 4 ' Causes InternetAutodial to fail if file and printer sharing is disabled for Microsoft® Windows® 95 or later.
Public Const INTERNET_DIAL_FORCE_PROMPT = &H2000       ' Ignores the "dial automatically" setting and forces the dialing user interface to be displayed.
Public Const INTERNET_DIAL_UNATTENDED = &H8000         ' Connects to the Internet through a modem, without displaying a user interface, if possible. Otherwise, the function will wait for user input.
Public Const INTERNET_DIAL_SHOW_OFFLINE = &H4000       ' Shows the Work Offline button instead of Cancel button in the dialing user interface.

' Constants - InternetCanonicalizeUrl.dwFlags / InternetCombineUrl.dwFlags
Public Const ICU_BROWSER_MODE = &H2000000       ' Does not encode or decode characters after "#" or "?", and does not remove trailing white space after "?". If this value is not specified, the entire URL is encoded and trailing white space is removed.
Public Const ICU_DECODE = &H10000000            ' Converts all %XX sequences to characters, including escape sequences, before the URL is parsed.
Public Const ICU_ENCODE_PERCENT = &H1000        ' Encodes any percent signs encountered. By default, percent signs are not encoded. This value is available in Microsoft® Internet Explorer 5 and later versions of the Microsoft® Win32® Internet functions.
Public Const ICU_ENCODE_SPACES_ONLY = &H4000000 ' Encodes spaces only.
Public Const ICU_NO_ENCODE = &H20000000         ' Does not convert unsafe characters to escape sequences.
Public Const ICU_NO_META = &H8000000            ' Does not remove meta sequences (such as "." and "..") from the URL.

' Constants - InternetConnect.nServerPort
Public Const INTERNET_DEFAULT_FTP_PORT = 21     ' Uses the default port for FTP servers (port 21).
Public Const INTERNET_DEFAULT_GOPHER_PORT = 70  ' Uses the default port for Gopher servers (port 70).
Public Const INTERNET_DEFAULT_HTTP_PORT = 80    ' Uses the default port for HTTP servers (port 80).
Public Const INTERNET_DEFAULT_HTTPS_PORT = 443  ' Uses the default port for HTTPS servers (port 443).
Public Const INTERNET_DEFAULT_SOCKS_PORT = 1080 ' Uses the default port for SOCKS firewall servers (port 1080).
Public Const INTERNET_INVALID_PORT_NUMBER = 0   ' Uses the default port for the service specified by dwService.

' Constants - InternetConnect.dwService
Public Const INTERNET_SERVICE_FTP = 1    ' FTP service.
Public Const INTERNET_SERVICE_GOPHER = 2 ' Gopher service.
Public Const INTERNET_SERVICE_HTTP = 3   ' HTTP service.

' Constants - InternetCrackUrl.dwFlags
'Public Const ICU_DECODE = &H10000000 'Converts encoded characters back to their normal form. This can be used only if the user provides buffers in the URL_COMPONENTS structure to copy the components into.
Public Const ICU_ESCAPE = &H80000000 'Converts all escape sequences (%xx) to their corresponding characters. This can be used only if the user provides buffers in the URL_COMPONENTS structure to copy the components into.


' Constants - InternetErrorDlg.dwFlags
Public Const FLAGS_ERROR_UI_FILTER_FOR_ERRORS = &H1    ' Scans the returned headers for errors. Call this flag after using HttpSendRequest. This option detects any hidden errors, such as an authentication error.
Public Const FLAGS_ERROR_UI_FLAGS_CHANGE_OPTIONS = &H2 ' If the function succeeds, stores the results of the dialog box in the Internet handle.
Public Const FLAGS_ERROR_UI_FLAGS_GENERATE_DATA = &H4  ' Queries the Internet handle for needed information. The function constructs the appropriate data structure for the error. (For example, for Cert CN failures, the function grabs the certificate.)
Public Const FLAGS_ERROR_UI_FLAGS_NO_UI = &H8          ' Undocumented
Public Const FLAGS_ERROR_UI_SERIALIZE_DIALOGS = &H10   ' Serializes authentication dialog boxes for concurrent requests on a password cache entry. The lppvData parameter should contain the address of a pointer to an INTERNET_AUTH_NOTIFY_DATA structure, and the client should implement a thread-safe, nonblocking callback function.

' Constants - InternetGetConnectedState.lpdwFlags / InternetGetConnectedStateEx.lpdwFlags
Public Const INTERNET_CONNECTION_CONFIGURED = &H40 'Local system has a valid connection to the Internet, but it may or may not be currently connected.
Public Const INTERNET_CONNECTION_LAN = &H2  'Local system uses a local area network to connect to the Internet.
Public Const INTERNET_CONNECTION_MODEM = &H1  'Local system uses a modem to connect to the Internet.
Public Const INTERNET_CONNECTION_MODEM_BUSY = &H8  'No longer used.
Public Const INTERNET_CONNECTION_OFFLINE = &H20 'Local system is in offline mode.
Public Const INTERNET_CONNECTION_PROXY = &H4  'Local system uses a proxy server to connect to the Internet.
Public Const INTERNET_RAS_INSTALLED = &H10 'Local system has RAS installed.

' Constants - InternetOpen.dwAccessType
Public Const INTERNET_OPEN_TYPE_DIRECT = 1    'Resolves all host names locally.
Public Const INTERNET_OPEN_TYPE_PRECONFIG = 0 'Retrieves the proxy or direct configuration from the registry.
Public Const INTERNET_OPEN_TYPE_PRECONFIG_WITH_NO_AUTOPROXY = 4 'Retrieves the proxy or direct configuration from the registry and prevents the use of a startup Microsoft® JScript® or Internet Setup (INS) file.
Public Const INTERNET_OPEN_TYPE_PROXY = 3     'Passes requests to the proxy unless a proxy bypass list is supplied and the name to be resolved bypasses the proxy. In this case, the function uses INTERNET_OPEN_TYPE_DIRECT.

' Constants - InternetOpen.dwFlags
Public Const INTERNET_FLAG_ASYNC = &H10000000     'Makes only asynchronous requests on handles descended from the handle returned from this function.
Public Const INTERNET_FLAG_FROM_CACHE = &H1000000 'Does not make network requests. All entities are returned from the cache. If the requested item is not in the cache, a suitable error, such as ERROR_FILE_NOT_FOUND, is returned.
Public Const INTERNET_FLAG_OFFLINE = INTERNET_FLAG_FROM_CACHE 'Identical to INTERNET_FLAG_FROM_CACHE. Does not make network requests. All entities are returned from the cache. If the requested item is not in the cache, a suitable error, such as ERROR_FILE_NOT_FOUND, is returned.

' Constants - InternetOpenUrl.dwFlags
Public Const INTERNET_FLAG_EXISTING_CONNECT = &H20000000      ' Attempts to use an existing InternetConnect object if one exists with the same attributes required to make the request. This is useful only with FTP operations, since FTP is the only protocol that typically performs multiple operations during the same session. The Microsoft® Win32® Internet API caches a single connection handle for each HINTERNET handle generated by InternetOpen.
'Public Const INTERNET_FLAG_HYPERLINK = &H400                 ' Forces a reload if there was no Expires time and no LastModified time returned from the server when determining whether to reload the item from the network.
Public Const INTERNET_FLAG_IGNORE_CERT_CN_INVALID = &H1000    ' Disables Win32 Internet function checking of SSL/PCT-based certificates that are returned from the server against the host name given in the request. Win32 Internet functions use a simple check against certificates by comparing for matching host names and simple wildcarding rules.
Public Const INTERNET_FLAG_IGNORE_CERT_DATE_INVALID = &H2000  ' Disables Win32 Internet function checking of SSL/PCT-based certificates for proper validity dates.
'Public Const INTERNET_FLAG_IGNORE_REDIRECT_TO_HTTP = &H8000  ' Disables the ability of the Win32 Internet functions to detect this special type of redirect. When this flag is used, Win32 Internet functions transparently allow redirects from HTTPS to HTTP URLs.
'Public Const INTERNET_FLAG_IGNORE_REDIRECT_TO_HTTPS = &H4000 ' Disables the ability of the Win32 Internet functions to detect this special type of redirect. When this flag is used, Win32 Internet functions transparently allow redirects from HTTP to HTTPS URLs.
'Public Const INTERNET_FLAG_KEEP_CONNECTION = &H400000        ' Uses keep-alive semantics, if available, for the connection. This flag is required for Microsoft Network (MSN), NT LAN Manager (NTLM), and other types of authentication.
'Public Const INTERNET_FLAG_NEED_FILE = &H10                  ' Causes a temporary file to be created if the file cannot be cached.
'Public Const INTERNET_FLAG_NO_AUTH = &H40000                 ' Does not attempt authentication automatically.
'Public Const INTERNET_FLAG_NO_AUTO_REDIRECT = &H200000       ' Does not automatically handle redirection in HttpSendRequest.
'Public Const INTERNET_FLAG_NO_CACHE_WRITE = &H4000000        ' Does not add the returned entity to the cache.
'Public Const INTERNET_FLAG_NO_COOKIES = &H80000              ' Does not automatically add cookie headers to requests, and does not automatically add returned cookies to the cookie database.
'Public Const INTERNET_FLAG_NO_UI = &H200                     ' Disables the cookie dialog box.
Public Const INTERNET_FLAG_PASSIVE = &H8000000                ' Uses passive FTP semantics. InternetOpenUrl uses this flag for FTP files and directories.
'Public Const INTERNET_FLAG_PRAGMA_NOCACHE = &H100            ' Forces the request to be resolved by the origin server, even if a cached copy exists on the proxy.
Public Const INTERNET_FLAG_RAW_DATA = &H40000000              ' Returns the data as a GOPHER_FIND_DATA structure when retrieving Gopher directory information, or as a WIN32_FIND_DATA  structure when retrieving FTP directory information. If this flag is not specified or if the call was made through a CERN proxy, InternetOpenUrl returns the HTML version of the directory.
'Public Const INTERNET_FLAG_RELOAD = &H80000000               ' Forces a download of the requested file, object, or directory listing from the origin server, not from the cache.
'Public Const INTERNET_FLAG_RESYNCHRONIZE = &H800             ' Reloads HTTP resources if the resource has been modified since the last time it was downloaded. All FTP and Gopher resources are reloaded.
'Public Const INTERNET_FLAG_SECURE = &H800000                 ' Uses secure transaction semantics. This translates to using Secure Sockets Layer/Private Communications Technology (SSL/PCT) and is only meaningful in HTTP requests.

' Constants - InternetReadFileEx.dwFlags
Public Const IRF_ASYNC = &H1       ' Undocumented
Public Const IRF_SYNC = &H4        ' Undocumented
Public Const IRF_USE_CONTEXT = &H8 ' Undocumented
Public Const IRF_NO_WAIT = &H8     ' Do not wait for data. If there is data available, the function returns either the amount of data requested or the amount of data available (whichever is smaller).

' Constants - InternetSetFilePointer.dwMoveMethod
Public Const FILE_BEGIN = 0   ' Starting point is zero or the beginning of the file. If FILE_BEGIN is specified, lDistanceToMove is interpreted as an unsigned location for the new file pointer.
Public Const FILE_CURRENT = 1 ' Current value of the file pointer is the starting point.
Public Const FILE_END = 2     ' Current end-of-file position is the starting point. This method fails if the content length is unknown.

' Constants - InternetSetOption.dwOption
Public Const INTERNET_OPTION_ASYNC = 30                       ' Not currently implemented.
Public Const INTERNET_OPTION_ASYNC_ID = 15                    ' Not implemented.
Public Const INTERNET_OPTION_ASYNC_PRIORITY = 16              ' Not currently implemented.
Public Const INTERNET_OPTION_BYPASS_EDITED_ENTRY = 64         ' Sets or retrieves the Boolean value that determines if the system should check the network for newer content and overwrite edited cache entries if a newer version is found. If set to TRUE, the system will check the network for newer content and overwrite the edited cache entry with the newer version. The default is FALSE, which indicates that the edited cache entry should be used without checking the network. This is used by InternetQueryOption and InternetSetOption. It is valid only in Microsoft® Internet Explorer 5 and later.
Public Const INTERNET_OPTION_CACHE_STREAM_HANDLE = 27         ' No longer supported.
Public Const INTERNET_OPTION_CACHE_TIMESTAMPS = 69            ' Retrieves an INTERNET_CACHE_TIMESTAMPS structure that contains the LastModified time and Expires time from the resource stored in the Internet cache. This value is used by InternetQueryOption.
Public Const INTERNET_OPTION_CALLBACK = 1                     ' Sets or retrieves the address of the callback function defined for this handle. This option can be used on all Appendix A: HINTERNET Handles handles. Used by InternetQueryOption and InternetSetOption.
Public Const INTERNET_OPTION_CALLBACK_FILTER = 54             ' Not currently implemented.
Public Const INTERNET_OPTION_CODEPAGE = 68                    ' Not currently implemented.
Public Const INTERNET_OPTION_CONNECT_BACKOFF = 4              ' Not currently implemented.
Public Const INTERNET_OPTION_CONNECT_LIMIT = 46               ' Not currently implemented.
Public Const INTERNET_OPTION_CONNECT_RETRIES = 3              ' Sets or retrieves an unsigned long integer value that contains the retry count to use for Internet connection requests. If a connection attempt still fails after the specified number of tries, the request is canceled. The default is five retries. This option can be used on any Appendix A: HINTERNET Handles handle, including a NULL handle. It is used by InternetQueryOption and InternetSetOption.
Public Const INTERNET_OPTION_CONNECT_TIME = 55                ' Not currently implemented.
Public Const INTERNET_OPTION_CONNECT_TIMEOUT = 2              ' Sets or retrieves an unsigned long integer value that contains the time-out value, in milliseconds, to use for Internet connection requests. If a connection request takes longer than this time-out value, the request is canceled. This option can be used on any Appendix A: HINTERNET Handles handle, including a NULL handle. It is used by InternetQueryOption and InternetSetOption.
Public Const INTERNET_OPTION_CONNECTED_STATE = 50             ' Sets or retrieves an unsigned long integer value that contains the connected state. This is used by InternetQueryOption and InternetSetOption.
Public Const INTERNET_OPTION_CONTEXT_VALUE = 45               ' Sets or retrieves a DWORD_PTR that contains the address of the context value associated with this Internet handle. This option can be used on any Appendix A: HINTERNET Handles handle. This is used by InternetQueryOption and InternetSetOption. Previously, this set the context value to the address stored in the DWORD(lpBuffer) pointer. This has been corrected so that the value stored in the buffer will be used and the INTERNET_OPTION_CONTEXT_VALUE flag will be assigned a new value. The old value, 10, has been preserved so that applications written for the old behavior are still supported.
Public Const INTERNET_OPTION_CONTROL_RECEIVE_TIMEOUT = 6      ' Identical to INTERNET_OPTION_RECEIVE_TIMEOUT. This is used by InternetQueryOption and InternetSetOption.
Public Const INTERNET_OPTION_CONTROL_SEND_TIMEOUT = 5         ' Identical to INTERNET_OPTION_SEND_TIMEOUT. This is used by InternetQueryOption and InternetSetOption.
Public Const INTERNET_OPTION_DATA_RECEIVE_TIMEOUT = 8         ' Not implemented.
Public Const INTERNET_OPTION_DATA_SEND_TIMEOUT = 7            ' Not implemented.
Public Const INTERNET_OPTION_DATAFILE_NAME = 33               ' Retrieves a string value that contains the name of the file backing a downloaded entity. This flag is valid after InternetOpenUrl, FtpOpenFile, GopherOpenFile, or HttpOpenRequest has completed. It is used by InternetQueryOption and InternetSetOption.
Public Const INTERNET_OPTION_DIGEST_AUTH_UNLOAD = 76          ' Causes the system to log off the Digest authentication SSPI package, purging all of the credentials created for the process. No buffer is required for this option. It is used by InternetSetOption.
Public Const INTERNET_OPTION_DISABLE_AUTODIAL = 70            ' Not currently implemented.
Public Const INTERNET_OPTION_DISCONNECTED_TIMEOUT = 49        ' Not currently implemented.
Public Const INTERNET_OPTION_END_BROWSER_SESSION = 42         ' Flushes entries not in use from the password cache on the hard drive. Also resets the cache time used when the synchronization mode is once-per-session. No buffer is required for this option. This is used by InternetSetOption.
Public Const INTERNET_OPTION_ERROR_MASK = 62                  ' Sets an unsigned long integer value that contains the error masks that can be handled by the client application. This can be a combination of the following values:
'                                                                INTERNET_ERROR_MASK_COMBINED_SEC_CERT                 - Indicates that the client application can handle security certificate error codes.
'                                                                INTERNET_ERROR_MASK_INSERT_CDROM                      - Indicates that the client application can handle the ERROR_INTERNET_INSERT_CDROM error code.
'                                                                INTERNET_ERROR_MASK_LOGIN_FAILURE_DISPLAY_ENTITY_BODY - Indicates that the client application can handle the ERROR_INTERNET_LOGIN_FAILURE_DISPLAY_ENTITY_BODY error code.
'                                                                INTERNET_ERROR_MASK_NEED_MSN_SSPI_PKG                 - Not currently implemented.
Public Const INTERNET_OPTION_EXTENDED_ERROR = 24              ' Retrieves an unsigned long integer value that contains a Microsoft® Windows® Sockets error code that was mapped to the ERROR_INTERNET_ error messages last returned in this thread context. This option is used on a NULL Appendix A: HINTERNET Handles handle by InternetQueryOption.
Public Const INTERNET_OPTION_FROM_CACHE_TIMEOUT = 63          ' Sets or retrieves an unsigned long integer value that contains the amount of time the system should wait for a response to a network request before checking the cache for a copy of the resource. If a network request takes longer than the time specified and the requested resource is available in the cache, the resource will be retrieved from the cache. This is used by InternetQueryOption and InternetSetOption.
Public Const INTERNET_OPTION_HANDLE_TYPE = 9                  ' Retrieves an unsigned long integer value that contains the type of the Internet handle passed in. This is used by InternetQueryOption on any Appendix A: HINTERNET Handles handle. Possible return values include:
'                                                                INTERNET_HANDLE_TYPE_CONNECT_FTP
'                                                                INTERNET_HANDLE_TYPE_CONNECT_GOPHER
'                                                                INTERNET_HANDLE_TYPE_CONNECT_HTTP
'                                                                INTERNET_HANDLE_TYPE_FILE_REQUEST
'                                                                INTERNET_HANDLE_TYPE_FTP_FILE
'                                                                INTERNET_HANDLE_TYPE_FTP_FILE_HTML
'                                                                INTERNET_HANDLE_TYPE_FTP_FIND
'                                                                INTERNET_HANDLE_TYPE_FTP_FIND_HTML
'                                                                INTERNET_HANDLE_TYPE_GOPHER_FILE
'                                                                INTERNET_HANDLE_TYPE_GOPHER_FILE_HTML
'                                                                INTERNET_HANDLE_TYPE_GOPHER_FIND
'                                                                INTERNET_HANDLE_TYPE_GOPHER_FIND_HTML
'                                                                INTERNET_HANDLE_TYPE_HTTP_REQUEST
'                                                                INTERNET_HANDLE_TYPE_INTERNET
Public Const INTERNET_OPTION_HTTP_VERSION = 59                ' Sets or retrieves an HTTP_VERSION_INFO structure that contains the HTTP version being supported. This must be used on a NULL handle. This is used by InternetQueryOption and InternetSetOption.
Public Const INTERNET_OPTION_IDLE_STATE = 51                  ' Not currently implemented.
Public Const INTERNET_OPTION_IGNORE_OFFLINE = 77              ' Sets or retrieves whether the global offline flag should be ignored. No buffer is required for this option. This is used by InternetQueryOption and InternetSetOption. This value was introduced in Internet Explorer 5.
Public Const INTERNET_OPTION_KEEP_CONNECTION = 22             ' Not currently implemented.
Public Const INTERNET_OPTION_LISTEN_TIMEOUT = 11              ' Not currently implemented.
Public Const INTERNET_OPTION_MAX_CONNS_PER_1_0_SERVER = 74    ' Sets or retrieves an unsigned long integer value that contains the maximum number of connections allowed per HTTP/1.0 server. This is used by InternetQueryOption and InternetSetOption. This value was introduced in Internet Explorer 5.
Public Const INTERNET_OPTION_MAX_CONNS_PER_SERVER = 73        ' Sets or retrieves an unsigned long integer value that contains the maximum number of connections allowed per server. This is used by InternetQueryOption and InternetSetOption. This value was introduced in Internet Explorer 5.
Public Const INTERNET_OPTION_OFFLINE_MODE = 26                ' Not currently implemented.
Public Const INTERNET_OPTION_OFFLINE_SEMANTICS = 52           ' Not currently implemented.
Public Const INTERNET_OPTION_PARENT_HANDLE = 21               ' Retrieves the parent handle to this handle. This option can be used on any Appendix A: HINTERNET Handles handle by InternetQueryOption.
Public Const INTERNET_OPTION_PASSWORD = 29                    ' Sets or retrieves a string value that contains the password associated with a handle returned by InternetConnect. This is used by InternetQueryOption and InternetSetOption.
Public Const INTERNET_OPTION_PER_CONNECTION_OPTION = 75       ' Sets or retrieves an INTERNET_PER_CONN_OPTION_LIST structure that specifies a list of options for a particular connection. This is used by InternetQueryOption and InternetSetOption. This option is only valid in Internet Explorer 5 and later.
Public Const INTERNET_OPTION_POLICY = 48                      ' Not currently implemented.
Public Const INTERNET_OPTION_PROXY = 38                       ' Sets or retrieves an INTERNET_PROXY_INFO structure that contains the proxy information on an existing InternetOpen handle when the Appendix A: HINTERNET Handles handle is not NULL. If the Appendix A: HINTERNET Handles handle is NULL, the function sets or queries the global proxy information. This option can be used on the Appendix A: HINTERNET Handles handle returned by InternetOpen. It is used by InternetQueryOption and InternetSetOption.
Public Const INTERNET_OPTION_PROXY_PASSWORD = 44              ' Sets or retrieves a string value that contains the password currently being used to access the proxy. This is used by InternetQueryOption and InternetSetOption.
Public Const INTERNET_OPTION_PROXY_USERNAME = 43              ' Sets or retrieves a string value that contains the user name currently being used to access the proxy. This is used by InternetQueryOption and InternetSetOption.
Public Const INTERNET_OPTION_READ_BUFFER_SIZE = 12            ' Sets or retrieves an unsigned long integer value that contains the size of the read buffer. This option can be used on Appendix A: HINTERNET Handles handles returned by FtpOpenFile, FtpFindFirstFile, and InternetConnect (FTP session only). This option is used by InternetQueryOption and InternetSetOption.
Public Const INTERNET_OPTION_RECEIVE_THROUGHPUT = 57          ' Not currently implemented.
Public Const INTERNET_OPTION_RECEIVE_TIMEOUT = 6              ' Sets or retrieves an unsigned long integer value that contains the time-out value, in milliseconds, to receive a response to a request. If the response takes longer than this time-out value, the request is canceled. This option can be used on any Appendix A: HINTERNET Handles handle, including a NULL handle. It is used by InternetQueryOption and InternetSetOption.
Public Const INTERNET_OPTION_REFRESH = 37                     ' Causes the proxy information to be reread from the registry for a handle. No buffer is required. This option can be used on the Appendix A: HINTERNET Handles handle returned by InternetOpen. It is used by InternetSetOption.
Public Const INTERNET_OPTION_REQUEST_FLAGS = 23               ' Retrieves an unsigned long integer value that contains the special status flags that indicate the status of the download currently in progress. This is used by InternetQueryOption. The INTERNET_OPTION_REQUEST_FLAGS option can be one of the following values:
'                                                                INTERNET_REQFLAG_ASYNC                - Not currently implemented.
'                                                                INTERNET_REQFLAG_CACHE_WRITE_DISABLED - Internet request cannot be cached (an HTTPS request, for example).
'                                                                INTERNET_REQFLAG_FROM_CACHE           - Response came from the cache.
'                                                                INTERNET_REQFLAG_NET_TIMEOUT          - Internet request timed out.
'                                                                INTERNET_REQFLAG_NO_HEADERS           - Original response contained no headers.
'                                                                INTERNET_REQFLAG_PASSIVE              - Not currently implemented.
'                                                                INTERNET_REQFLAG_VIA_PROXY            - Request was made through a proxy.
Public Const INTERNET_OPTION_REQUEST_PRIORITY = 58            ' Sets or retrieves an unsigned long integer value that contains the priority of requests competing for a connection on an HTTP handle. This is used by InternetQueryOption and InternetSetOption.
Public Const INTERNET_OPTION_RESET_URLCACHE_SESSION = 60      ' Starts a new cache session for the process. No buffer is required. This is used by InternetSetOption.
Public Const INTERNET_OPTION_SECONDARY_CACHE_KEY = 53         ' Sets or retrieves a string value that contains the secondary cache key. This is used by InternetQueryOption and InternetSetOption.
Public Const INTERNET_OPTION_SECURITY_CERTIFICATE = 35        ' Retrieves the certificate for an SSL/PCT (Secure Sockets Layer/Private Communications Technology) server into a formatted string. This is used by InternetQueryOption.
Public Const INTERNET_OPTION_SECURITY_CERTIFICATE_STRUCT = 32 ' Retrieves the certificate for an SSL/PCT server into the INTERNET_CERTIFICATE_INFO structure. This is used by InternetQueryOption.
Public Const INTERNET_OPTION_SECURITY_FLAGS = 31              ' Retrieves an unsigned long integer value that contains the security flags for a handle. This option is used by InternetQueryOption. It can be a combination of these values:
'                                                                SECURITY_FLAG_128BIT                   - Identical to the preferred value SECURITY_FLAG_STRENGTH_STRONG. This is only returned in a call to InternetQueryOption.
'                                                                SECURITY_FLAG_40BIT                    - Identical to the preferred value SECURITY_FLAG_STRENGTH_WEAK. This is only returned in a call to InternetQueryOption.
'                                                                SECURITY_FLAG_56BIT                    - Identical to the preferred value SECURITY_FLAG_STRENGTH_MEDIUM. This is only returned in a call to InternetQueryOption.
'                                                                SECURITY_FLAG_FORTEZZA                 - Indicates Fortezza has been used to provide secrecy, authentication, and/or integrity for the specified connection.
'                                                                SECURITY_FLAG_IETFSSL4                 - Not currently implemented.
'                                                                SECURITY_FLAG_IGNORE_CERT_CN_INVALID   - Ignores the ERROR_INTERNET_SEC_CERT_CN_INVALID error message.
'                                                                SECURITY_FLAG_IGNORE_CERT_DATE_INVALID - Ignores the ERROR_INTERNET_SEC_CERT_DATE_INVALID error message.
'                                                                SECURITY_FLAG_IGNORE_REDIRECT_TO_HTTP  - Ignores the ERROR_INTERNET_HTTPS_TO_HTTP_ON_REDIR error message.
'                                                                SECURITY_FLAG_IGNORE_REDIRECT_TO_HTTPS - Ignores the ERROR_INTERNET_HTTP_TO_HTTPS_ON_REDIR error message.
'                                                                SECURITY_FLAG_IGNORE_REVOCATION        - Ignores certificate revocation problems.
'                                                                SECURITY_FLAG_IGNORE_UNKNOWN_CA        - Ignores unknown certificate authority problems.
'                                                                SECURITY_FLAG_IGNORE_WRONG_USAGE       - Ignores incorrect usage problems.
'                                                                SECURITY_FLAG_NORMALBITNESS            - Identical to the value SECURITY_FLAG_STRENGTH_WEAK. This is only returned in a call to InternetQueryOption.
'                                                                SECURITY_FLAG_PCT                      - Not currently implemented.
'                                                                SECURITY_FLAG_PCT4                     - Not currently implemented.
'                                                                SECURITY_FLAG_SECURE                   - Uses secure transfers. This is only returned in a call to InternetQueryOption.
'                                                                SECURITY_FLAG_SSL                      - Not currently implemented.
'                                                                SECURITY_FLAG_SSL3                     - Not currently implemented.
'                                                                SECURITY_FLAG_STRENGTH_MEDIUM          - Uses medium (56-bit) encryption. This is only returned in a call to InternetQueryOption.
'                                                                SECURITY_FLAG_STRENGTH_STRONG          - Uses strong (128-bit) encryption. This is only returned in a call to InternetQueryOption.
'                                                                SECURITY_FLAG_STRENGTH_WEAK            - Uses weak (40-bit) encryption. This is only returned in a call to InternetQueryOption.
'                                                                SECURITY_FLAG_UNKNOWNBIT               - The bit size used in the encryption is unknown. This is only returned in a call to InternetQueryOption.
Public Const INTERNET_OPTION_SECURITY_KEY_BITNESS = 36      ' Retrieves an unsigned long integer value that contains the bit size of the encryption key. The larger the number, the greater the encryption strength being used. This is used by InternetQueryOption.
Public Const INTERNET_OPTION_SEND_THROUGHPUT = 56           ' Not currently implemented.
Public Const INTERNET_OPTION_SEND_TIMEOUT = 5               ' Sets or retrieves an unsigned long integer value that contains the time-out value, in milliseconds, to send a request. If the send takes longer than this time-out value, the send is canceled. This option can be used on any Appendix A: HINTERNET Handles handle, including a NULL handle. It is used by InternetQueryOption and InternetSetOption.
Public Const INTERNET_OPTION_SETTINGS_CHANGED = 39          ' Informs the system that the registry settings have been changed so that it will check the settings on the next call to InternetConnect. This is used by InternetSetOption.
Public Const INTERNET_OPTION_URL = 34                       ' Retrieves a string value that contains the full URL of a downloaded resource. If the original URL contained any extra information (such as search strings or anchors), or if the call was redirected, the URL returned will differ from the original. This option is valid on Appendix A: HINTERNET Handles handles returned by InternetOpenUrl, FtpOpenFile, GopherOpenFile, or HttpOpenRequest. It is used by InternetQueryOption.
Public Const INTERNET_OPTION_USER_AGENT = 41                ' Sets or retrieves the user agent string on handles supplied by InternetOpen and used in subsequent HttpSendRequest functions, as long as it is not overridden by a header added by HttpAddRequestHeaders or HttpSendRequest. This is used by InternetQueryOption and InternetSetOption.
Public Const INTERNET_OPTION_USERNAME = 28                  ' Sets or retrieves a string that contains the user name associated with a handle returned by InternetConnect. This is used by InternetQueryOption and InternetSetOption.
Public Const INTERNET_OPTION_VERSION = 40                   ' Retrieves an INTERNET_VERSION_INFO structure that contains the version number of Wininet.dll. This option can be used on a NULL Appendix A: HINTERNET Handles handle by InternetQueryOption.
Public Const INTERNET_OPTION_WRITE_BUFFER_SIZE = 13         ' Sets or retrieves an unsigned long integer value that contains the size of the write buffer. This option can be used on Appendix A: HINTERNET Handles handles returned by FtpOpenFile and InternetConnect (FTP session only). It is used by InternetQueryOption and InternetSetOption.
'Public Const INTERNET_OPTION_CLIENT_CERT_CONTEXT = ?       ' This flag is not supported by InternetQueryOption. The LPVOID(lpBuffer) parameter must be a pointer to a CERT CONTEXT  structure and not a pointer to a CERT CONTEXT pointer. If an application receives ERROR_INTERNET_CLIENT_AUTH_CERT_NEEDED, it must call InternetErrorDlg or use InternetSetOption to supply a certificate before retrying the request. CertDuplicateCertificateContext is then called so that the certificate context passed can be independently released by the application.

' Constants - CallbackProc.dwInternetStatus
Private Const INTERNET_STATUS_CLOSING_CONNECTION = 50     ' Closing the connection to the server. The lpvStatusInformation parameter is NULL.
Private Const INTERNET_STATUS_CONNECTED_TO_SERVER = 21    ' Successfully connected to the socket address (SOCKADDR) pointed to by lpvStatusInformation.
Private Const INTERNET_STATUS_CONNECTING_TO_SERVER = 20   ' Connecting to the socket address (SOCKADDR) pointed to by lpvStatusInformation.
Private Const INTERNET_STATUS_CONNECTION_CLOSED = 51      ' Successfully closed the connection to the server. The lpvStatusInformation parameter is NULL.
Private Const INTERNET_STATUS_CTL_RESPONSE_RECEIVED = 42  ' Not implemented
Private Const INTERNET_STATUS_DETECTING_PROXY = 80        ' Notifies the client application that a proxy has been detected.
Private Const INTERNET_STATUS_HANDLE_CLOSING = 70         ' This handle value has been terminated.
Private Const INTERNET_STATUS_HANDLE_CREATED = 60         ' Used by InternetConnect to indicate it has created the new handle. This lets the application call InternetCloseHandle from another thread, if the connect is taking too long. The lpvStatusInformation parameter contains the address of an INTERNET_ASYNC_RESULT structure.
Private Const INTERNET_STATUS_INTERMEDIATE_RESPONSE = 120 ' Received an intermediate (100 level) status code message from the server.
Private Const INTERNET_STATUS_NAME_RESOLVED = 11          ' Successfully found the IP address of the name contained in lpvStatusInformation.
Private Const INTERNET_STATUS_PREFETCH = 43               ' Not implemented
Private Const INTERNET_STATUS_RECEIVING_RESPONSE = 40     ' Waiting for the server to respond to a request. The lpvStatusInformation parameter is NULL.
Private Const INTERNET_STATUS_REDIRECT = 110              ' An HTTP request is about to automatically redirect the request. The lpvStatusInformation parameter points to the new URL. At this point, the application can read any data returned by the server with the redirect response and can query the response headers. It can also cancel the operation by closing the handle. This callback is not made if the original request specified INTERNET_FLAG_NO_AUTO_REDIRECT.
Private Const INTERNET_STATUS_REQUEST_COMPLETE = 100      ' An asynchronous operation has been completed. The lpvStatusInformation parameter contains the address of an INTERNET_ASYNC_RESULT structure.
Private Const INTERNET_STATUS_REQUEST_SENT = 31           ' Successfully sent the information request to the server. The lpvStatusInformation parameter points to a DWORD containing the number of bytes sent.
Private Const INTERNET_STATUS_RESOLVING_NAME = 10         ' Looking up the IP address of the name contained in lpvStatusInformation.
Private Const INTERNET_STATUS_RESPONSE_RECEIVED = 41      ' Successfully received a response from the server. The lpvStatusInformation parameter points to a DWORD containing the number of bytes received.
Private Const INTERNET_STATUS_SENDING_REQUEST = 30        ' Sending the information request to the server. The lpvStatusInformation parameter is NULL.
Private Const INTERNET_STATUS_STATE_CHANGE = 200          ' Moved between a secure (HTTPS) and a nonsecure (HTTP) site. This can be one of the following values:
Private Const INTERNET_STATE_CONNECTED = &H1              ' Connected state (mutually exclusive with disconnected state).
Private Const INTERNET_STATE_DISCONNECTED = &H2           ' Disconnected state. No network connection could be established.
Private Const INTERNET_STATE_DISCONNECTED_BY_USER = &H10  ' Disconnected by user request.
Private Const INTERNET_STATE_IDLE = &H100                 ' No network requests are being made by the Microsoft® Win32® Internet functions.
Private Const INTERNET_STATE_BUSY = &H200                 ' Network requests are being made by the Win32 Internet functions.
Private Const INTERNET_STATUS_USER_INPUT_REQUIRED = 140   ' The request requires user input to be completed.

' Constants - InternetTimeFromSystemTime.dwRFC
Public Const INTERNET_RFC1123_FORMAT = 0

' Constants - SetUrlCacheEntryGroup.dwFlags
Public Const INTERNET_CACHE_GROUP_ADD = 0    ' Adds the cache entry to the cache group.
Public Const INTERNET_CACHE_GROUP_REMOVE = 0 ' Removes the entry from the cache group.

' Constants - SetUrlCacheEntryInfo.dwFieldControl
Public Const CACHE_ENTRY_ACCTIME_FC = &H100      ' Sets the last access time.
Public Const CACHE_ENTRY_ATTRIBUTE_FC = &H4      ' Sets the cache entry type.
Public Const CACHE_ENTRY_EXEMPT_DELTA_FC = &H800 ' Sets the exempt delta.
Public Const CACHE_ENTRY_EXPTIME_FC = &H80       ' Sets the expire time.
Public Const CACHE_ENTRY_HEADERINFO_FC = &H400   ' Not currently implemented.
Public Const CACHE_ENTRY_HITRATE_FC = &H10       ' Sets the hit rate.
Public Const CACHE_ENTRY_MODTIME_FC = &H40       ' Sets the last modified time.
Public Const CACHE_ENTRY_SYNCTIME_FC = &H200     ' Sets the last sync time.

' Internet API Error Returns
Public Const INTERNET_ERROR_BASE = 12000
Public Const INTERNET_ERROR_FIRST = (INTERNET_ERROR_BASE + 1)
Public Const INTERNET_ERROR_LAST = (INTERNET_ERROR_BASE + 174)

Public Const ERROR_INTERNET_OUT_OF_HANDLES = (INTERNET_ERROR_BASE + 1) 'No more handles could be generated at this time.
Public Const ERROR_INTERNET_TIMEOUT = (INTERNET_ERROR_BASE + 2) 'The request has timed out.
Public Const ERROR_INTERNET_EXTENDED_ERROR = (INTERNET_ERROR_BASE + 3) 'An extended error was returned from the server. This is typically a string or buffer containing a verbose error message. Call InternetGetLastResponseInfo to retrieve the error text.
Public Const ERROR_INTERNET_INTERNAL_ERROR = (INTERNET_ERROR_BASE + 4) 'An internal error has occurred.
Public Const ERROR_INTERNET_INVALID_URL = (INTERNET_ERROR_BASE + 5) 'The URL is invalid.
Public Const ERROR_INTERNET_UNRECOGNIZED_SCHEME = (INTERNET_ERROR_BASE + 6) 'The URL scheme could not be recognized, or is not supported.
Public Const ERROR_INTERNET_NAME_NOT_RESOLVED = (INTERNET_ERROR_BASE + 7) 'The server name could not be resolved.
Public Const ERROR_INTERNET_PROTOCOL_NOT_FOUND = (INTERNET_ERROR_BASE + 8) 'The requested protocol could not be located.
Public Const ERROR_INTERNET_INVALID_OPTION = (INTERNET_ERROR_BASE + 9) 'A request to InternetQueryOption or InternetSetOption specified an invalid option value.
Public Const ERROR_INTERNET_BAD_OPTION_LENGTH = (INTERNET_ERROR_BASE + 10) 'The length of an option supplied to InternetQueryOption or InternetSetOption is incorrect for the type of option specified.
Public Const ERROR_INTERNET_OPTION_NOT_SETTABLE = (INTERNET_ERROR_BASE + 11) 'The requested option cannot be set, only queried.
Public Const ERROR_INTERNET_SHUTDOWN = (INTERNET_ERROR_BASE + 12) 'The Win32 Internet function support is being shut down or unloaded.
Public Const ERROR_INTERNET_INCORRECT_USER_NAME = (INTERNET_ERROR_BASE + 13) 'The request to connect and log on to an FTP server could not be completed because the supplied user name is incorrect.
Public Const ERROR_INTERNET_INCORRECT_PASSWORD = (INTERNET_ERROR_BASE + 14) 'The request to connect and log on to an FTP server could not be completed because the supplied password is incorrect.
Public Const ERROR_INTERNET_LOGIN_FAILURE = (INTERNET_ERROR_BASE + 15) 'The request to connect and log on to an FTP server failed.
Public Const ERROR_INTERNET_INVALID_OPERATION = (INTERNET_ERROR_BASE + 16) 'The requested operation is invalid.
Public Const ERROR_INTERNET_OPERATION_CANCELLED = (INTERNET_ERROR_BASE + 17) 'The operation was canceled, usually because the handle on which the request was operating was closed before the operation completed.
Public Const ERROR_INTERNET_INCORRECT_HANDLE_TYPE = (INTERNET_ERROR_BASE + 18) 'The type of handle supplied is incorrect for this operation.
Public Const ERROR_INTERNET_INCORRECT_HANDLE_STATE = (INTERNET_ERROR_BASE + 19) 'The requested operation cannot be carried out because the handle supplied is not in the correct state.
Public Const ERROR_INTERNET_NOT_PROXY_REQUEST = (INTERNET_ERROR_BASE + 20) 'The request cannot be made via a proxy.
Public Const ERROR_INTERNET_REGISTRY_VALUE_NOT_FOUND = (INTERNET_ERROR_BASE + 21) 'A required registry value could not be located.
Public Const ERROR_INTERNET_BAD_REGISTRY_PARAMETER = (INTERNET_ERROR_BASE + 22) 'A required registry value was located but is an incorrect type or has an invalid value.
Public Const ERROR_INTERNET_NO_DIRECT_ACCESS = (INTERNET_ERROR_BASE + 23) 'Direct network access cannot be made at this time.
Public Const ERROR_INTERNET_NO_CONTEXT = (INTERNET_ERROR_BASE + 24) 'An asynchronous request could not be made because a zero context value was supplied.
Public Const ERROR_INTERNET_NO_CALLBACK = (INTERNET_ERROR_BASE + 25) 'An asynchronous request could not be made because a callback function has not been set.
Public Const ERROR_INTERNET_REQUEST_PENDING = (INTERNET_ERROR_BASE + 26) 'The required operation could not be completed because one or more requests are pending.
Public Const ERROR_INTERNET_INCORRECT_FORMAT = (INTERNET_ERROR_BASE + 27) 'The format of the request is invalid.
Public Const ERROR_INTERNET_ITEM_NOT_FOUND = (INTERNET_ERROR_BASE + 28) 'The requested item could not be located.
Public Const ERROR_INTERNET_CANNOT_CONNECT = (INTERNET_ERROR_BASE + 29) 'The attempt to connect to the server failed.
Public Const ERROR_INTERNET_CONNECTION_ABORTED = (INTERNET_ERROR_BASE + 30) 'The connection with the server has been terminated.
Public Const ERROR_INTERNET_CONNECTION_RESET = (INTERNET_ERROR_BASE + 31) 'The connection with the server has been reset.
Public Const ERROR_INTERNET_FORCE_RETRY = (INTERNET_ERROR_BASE + 32) 'The Win32 Internet function needs to redo the request.
Public Const ERROR_INTERNET_INVALID_PROXY_REQUEST = (INTERNET_ERROR_BASE + 33) 'The request to the proxy was invalid.
Public Const ERROR_INTERNET_NEED_UI = (INTERNET_ERROR_BASE + 34) 'A user interface or other blocking operation has been requested.
Public Const ERROR_INTERNET_HANDLE_EXISTS = (INTERNET_ERROR_BASE + 36) 'The request failed because the handle already exists.
Public Const ERROR_INTERNET_SEC_CERT_DATE_INVALID = (INTERNET_ERROR_BASE + 37) 'SSL certificate date that was received from the server is bad. The certificate is expired.
Public Const ERROR_INTERNET_SEC_CERT_CN_INVALID = (INTERNET_ERROR_BASE + 38) 'SSL certificate common name (host name field) is incorrect—for example, if you entered www.server.com and the common name on the certificate says www.different.com.
Public Const ERROR_INTERNET_HTTP_TO_HTTPS_ON_REDIR = (INTERNET_ERROR_BASE + 39) 'The application is moving from a non-SSL to an SSL connection because of a redirect.
Public Const ERROR_INTERNET_HTTPS_TO_HTTP_ON_REDIR = (INTERNET_ERROR_BASE + 40) 'The application is moving from an SSL to an non-SSL connection because of a redirect.
Public Const ERROR_INTERNET_MIXED_SECURITY = (INTERNET_ERROR_BASE + 41) 'The content is not entirely secure. Some of the content being viewed may have come from unsecured servers.
Public Const ERROR_INTERNET_CHG_POST_IS_NON_SECURE = (INTERNET_ERROR_BASE + 42) 'The application is posting and attempting to change multiple lines of text on a server that is not secure.
Public Const ERROR_INTERNET_POST_IS_NON_SECURE = (INTERNET_ERROR_BASE + 43) 'The application is posting data to a sever that is not secure.
Public Const ERROR_INTERNET_CLIENT_AUTH_CERT_NEEDED = (INTERNET_ERROR_BASE + 44) 'The server is requesting client authentication.
Public Const ERROR_INTERNET_INVALID_CA = (INTERNET_ERROR_BASE + 45) 'The function is unfamiliar with the Certificate Authority that generated the server's certificate.
Public Const ERROR_INTERNET_CLIENT_AUTH_NOT_SETUP = (INTERNET_ERROR_BASE + 46) 'Client authorization is not set up on this computer.
Public Const ERROR_INTERNET_ASYNC_THREAD_FAILED = (INTERNET_ERROR_BASE + 47) 'The application could not start an asynchronous thread.
Public Const ERROR_INTERNET_REDIRECT_SCHEME_CHANGE = (INTERNET_ERROR_BASE + 48) 'The function could not handle the redirection, because the scheme changed (for example, HTTP to FTP).
Public Const ERROR_INTERNET_DIALOG_PENDING = (INTERNET_ERROR_BASE + 49) 'Another thread has a password dialog box in progress.
Public Const ERROR_INTERNET_RETRY_DIALOG = (INTERNET_ERROR_BASE + 50) 'The dialog box should be retried.
Public Const ERROR_INTERNET_HTTPS_HTTP_SUBMIT_REDIR = (INTERNET_ERROR_BASE + 52) 'The data being submitted to an SSL connection is being redirected to a non-SSL connection.
Public Const ERROR_INTERNET_INSERT_CDROM = (INTERNET_ERROR_BASE + 53) 'The request requires a CD-ROM to be inserted in the CD-ROM drive to locate the resource requested.
Public Const ERROR_INTERNET_FORTEZZA_LOGIN_NEEDED = (INTERNET_ERROR_BASE + 54) 'The requested resource requires Fortezza authentication.
Public Const ERROR_INTERNET_SEC_CERT_ERRORS = (INTERNET_ERROR_BASE + 55) 'The SSL certificate contains errors.
Public Const ERROR_INTERNET_SEC_CERT_NO_REV = (INTERNET_ERROR_BASE + 56)
Public Const ERROR_INTERNET_SEC_CERT_REV_FAILED = (INTERNET_ERROR_BASE + 57)

' FTP API Errors
Public Const ERROR_FTP_TRANSFER_IN_PROGRESS = (INTERNET_ERROR_BASE + 110) 'The requested operation cannot be made on the FTP session handle because an operation is already in progress.
Public Const ERROR_FTP_DROPPED = (INTERNET_ERROR_BASE + 111) 'The FTP operation was not completed because the session was aborted.
Public Const ERROR_FTP_NO_PASSIVE_MODE = (INTERNET_ERROR_BASE + 112) 'Passive mode is not available on the server.

' Gopher API Errors
Public Const ERROR_GOPHER_PROTOCOL_ERROR = (INTERNET_ERROR_BASE + 130) 'An error was detected while parsing data returned from the Gopher server.
Public Const ERROR_GOPHER_NOT_FILE = (INTERNET_ERROR_BASE + 131) 'The request must be made for a file locator.
Public Const ERROR_GOPHER_DATA_ERROR = (INTERNET_ERROR_BASE + 132) 'An error was detected while receiving data from the Gopher server.
Public Const ERROR_GOPHER_END_OF_DATA = (INTERNET_ERROR_BASE + 133) 'The end of the data has been reached.
Public Const ERROR_GOPHER_INVALID_LOCATOR = (INTERNET_ERROR_BASE + 134) 'The supplied locator is not valid.
Public Const ERROR_GOPHER_INCORRECT_LOCATOR_TYPE = (INTERNET_ERROR_BASE + 135) 'The type of the locator is not correct for this operation.
Public Const ERROR_GOPHER_NOT_GOPHER_PLUS = (INTERNET_ERROR_BASE + 136) 'The requested operation can be made only against a Gopher+ server, or with a locator that specifies a Gopher+ operation.
Public Const ERROR_GOPHER_ATTRIBUTE_NOT_FOUND = (INTERNET_ERROR_BASE + 137) 'The requested attribute could not be located.
Public Const ERROR_GOPHER_UNKNOWN_LOCATOR = (INTERNET_ERROR_BASE + 138) 'The locator type is unknown.

' HTTP API Errors
Public Const ERROR_HTTP_HEADER_NOT_FOUND = (INTERNET_ERROR_BASE + 150) 'The requested header could not be located.
Public Const ERROR_HTTP_DOWNLEVEL_SERVER = (INTERNET_ERROR_BASE + 151) 'The server did not return any headers.
Public Const ERROR_HTTP_INVALID_SERVER_RESPONSE = (INTERNET_ERROR_BASE + 152) 'The server response could not be parsed.
Public Const ERROR_HTTP_INVALID_HEADER = (INTERNET_ERROR_BASE + 153) 'The supplied header is invalid.
Public Const ERROR_HTTP_INVALID_QUERY_REQUEST = (INTERNET_ERROR_BASE + 154) 'The request made to HttpQueryInfo is invalid.
Public Const ERROR_HTTP_HEADER_ALREADY_EXISTS = (INTERNET_ERROR_BASE + 155) 'The header could not be added because it already exists.
Public Const ERROR_HTTP_REDIRECT_FAILED = (INTERNET_ERROR_BASE + 156) 'The redirection failed because either the scheme changed (for example, HTTP to FTP) or all attempts made to redirect failed (default is five attempts).
Public Const ERROR_HTTP_NOT_REDIRECTED = (INTERNET_ERROR_BASE + 160) 'The HTTP request was not redirected.
Public Const ERROR_HTTP_COOKIE_NEEDS_CONFIRMATION = (INTERNET_ERROR_BASE + 161) 'The HTTP cookie requires confirmation.
Public Const ERROR_HTTP_COOKIE_DECLINED = (INTERNET_ERROR_BASE + 162) 'The HTTP cookie was declined by the server.
Public Const ERROR_HTTP_REDIRECT_NEEDS_CONFIRMATION = (INTERNET_ERROR_BASE + 168) 'The redirection requires user confirmation.

' Additional Internet API Error Codes
Public Const ERROR_INTERNET_SECURITY_CHANNEL_ERROR = (INTERNET_ERROR_BASE + 157) 'The application experienced an internal error loading the SSL libraries.
Public Const ERROR_INTERNET_UNABLE_TO_CACHE_FILE = (INTERNET_ERROR_BASE + 158) 'The function was unable to cache the file.
Public Const ERROR_INTERNET_TCPIP_NOT_INSTALLED = (INTERNET_ERROR_BASE + 159) 'The required protocol stack is not loaded and the application cannot start WinSock.
Public Const ERROR_INTERNET_DISCONNECTED = (INTERNET_ERROR_BASE + 163) 'The Internet connection has been lost.
Public Const ERROR_INTERNET_SERVER_UNREACHABLE = (INTERNET_ERROR_BASE + 164) 'The Web site or server indicated is unreachable.
Public Const ERROR_INTERNET_PROXY_SERVER_UNREACHABLE = (INTERNET_ERROR_BASE + 165) 'The designated proxy server cannot be reached.
Public Const ERROR_INTERNET_BAD_AUTO_PROXY_SCRIPT = (INTERNET_ERROR_BASE + 166) 'There was an error in the automatic proxy configuration script.
Public Const ERROR_INTERNET_UNABLE_TO_DOWNLOAD_SCRIPT = (INTERNET_ERROR_BASE + 167) 'The automatic proxy configuration script could not be downloaded. The INTERNET_FLAG_MUST_CACHE_REQUEST flag was set.
Public Const ERROR_INTERNET_SEC_INVALID_CERT = (INTERNET_ERROR_BASE + 169) 'SSL certificate is invalid.
Public Const ERROR_INTERNET_SEC_CERT_REVOKED = (INTERNET_ERROR_BASE + 170) 'SSL certificate was revoked.

' InternetAutodial Specific Errors
Public Const ERROR_INTERNET_FAILED_DUETOSECURITYCHECK = (INTERNET_ERROR_BASE + 171) 'The function failed due to a security check.
Public Const ERROR_INTERNET_NOT_INITIALIZED = (INTERNET_ERROR_BASE + 172) 'Initialization of the Win32 Internet API has not occurred. Indicates that a higher-level function, such as InternetOpen, has not been called yet.
Public Const ERROR_INTERNET_NEED_MSN_SSPI_PKG = (INTERNET_ERROR_BASE + 173) 'Not currently implemented.
Public Const ERROR_INTERNET_LOGIN_FAILURE_DISPLAY_ENTITY_BODY = (INTERNET_ERROR_BASE + 174) 'The MS-Logoff digest header has been returned from the Web site. This header specifically instructs the digest package to purge credentials for the associated realm. This error will only be returned if INTERNET_ERROR_MASK_LOGIN_FAILURE_DISPLAY_ENTITY_BODY has been set.

' InternetDial Specific Errors
Public Const RASBASE = 600
Public Const ERROR_NO_CONNECTION = (RASBASE + 68)
Public Const ERROR_USER_DISCONNECTION = (RASBASE + 31)

' Internet Related Win32 Errors
Public Const ERROR_SUCCESS = 0 ' Process completed successfully
Public Const ERROR_INVALID_HANDLE = 6 'The handle that was passed to the API has been either invalidated or closed. (Win32 error code)
Public Const ERROR_NO_MORE_FILES = 18  'No more files have been found. (Win32 error code)
Public Const ERROR_NO_MORE_ITEMS = 259 'No more items have been found. (Win32 error code)
Public Const ERROR_BAD_PATHNAME = 161
Public Const ERROR_INSUFFICIENT_BUFFER = 122
Public Const ERROR_INVALID_PARAMETER = 87
Public Const ERROR_DISK_FULL = 112
Public Const ERROR_FILE_NOT_FOUND = 2
Public Const ERROR_ACCESS_DENIED = 5
Public Const ERROR_CANCELLED = 1223
Public Const ERROR_NOT_ENOUGH_MEMORY = 8

' Constants - Win32 API Related
Public Const TIME_DAY_ZERO         As Double = 109205# ' Abs(CDbl(#01-01-1601#))
Public Const TIME_MILLISEC_PER_DAY As Double = 10000000# * 60# * 60# * 24# / 10000#

Public Const FORMAT_MESSAGE_ALLOCATE_BUFFER = &H100 ' Specifies that the lpBuffer parameter is a pointer to a PVOID pointer, and that the nSize parameter specifies the minimum number of TCHARs to allocate for an output message buffer. The function allocates a buffer large enough to hold the formatted message, and places a pointer to the allocated buffer at the address specified by lpBuffer. The caller should use the LocalFree function to free the buffer when it is no longer needed.
Public Const FORMAT_MESSAGE_IGNORE_INSERTS = &H200  ' Specifies that insert sequences in the message definition are to be ignored and passed through to the output buffer unchanged. This flag is useful for fetching a message for later formatting. If this flag is set, the Arguments parameter is ignored.
Public Const FORMAT_MESSAGE_FROM_STRING = &H400     ' Specifies that lpSource is a pointer to a null-terminated message definition. The message definition may contain insert sequences, just as the message text in a message table resource may. Cannot be used with FORMAT_MESSAGE_FROM_HMODULE or FORMAT_MESSAGE_FROM_SYSTEM.
Public Const FORMAT_MESSAGE_FROM_HMODULE = &H800    ' Specifies that lpSource is a module handle containing the message-table resource(s) to search. If this lpSource handle is NULL, the current process's application image file will be searched. Cannot be used with FORMAT_MESSAGE_FROM_STRING.
Public Const FORMAT_MESSAGE_FROM_SYSTEM = &H1000    ' Specifies that the function should search the system message-table resource(s) for the requested message. If this flag is specified with FORMAT_MESSAGE_FROM_HMODULE, the function searches the system message table if the message is not found in the module specified by lpSource. Cannot be used with FORMAT_MESSAGE_FROM_STRING.  If this flag is specified, an application can pass the result of the GetLastError function to retrieve the message text for a system-defined error.
Public Const FORMAT_MESSAGE_ARGUMENT_ARRAY = &H2000 ' Specifies that the Arguments parameter is not a va_list structure, but instead is just a pointer to an array of values that represent the arguments.

' Variables
Private PrevCallbackAddr As Long

' Win32 API Declarations
Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Public Declare Sub SetLastError Lib "kernel32" (ByVal dwErrCode As Long)
Public Declare Function FileTimeToLocalFileTime Lib "kernel32" (ByRef lpFileTime As Currency, ByRef lpLocalFileTime As Currency) As Long
Public Declare Function GetLastError Lib "kernel32" () As Long
Public Declare Function FormatMessage Lib "kernel32" Alias "FormatMessageA" (ByVal dwFlags As Long, ByRef lpSource As Any, ByVal dwMessageId As Long, ByVal dwLanguageId As Long, ByVal lpBuffer As String, ByVal nSize As Long, ByRef Arguments As Long) As Long
Public Declare Function StringFromPointer Lib "kernel32" Alias "lstrcpy" (ByVal lpDestinationStr As String, ByVal lpStringPointer As Long) As Long




'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX




'=============================================================================================================
' CommitUrlCacheEntry
'
' Minimum Availability : Internet Explorer 3.0
'
' Parameters:
' ¯¯¯¯¯¯¯¯¯¯¯
' lpszUrlName       [in] Address of a string variable that contains the source name of the cache entry. The name string must be unique and should not contain any escape characters.
' lpszLocalFileName [in] Address of a string variable that contains the name of the local file that is being cached. This should be the same name as that returned by CreateUrlCacheEntry.
' ExpireTime        [in] FILETIME  structure that contains the expire date and time (in Greenwich mean time) of the file that is being cached. If the expire date and time is unknown, set this parameter to zero.
' LastModifiedTime  [in] FILETIME structure that contains the last modified date and time (in Greenwich mean time) of the URL that is being cached. If the last modified date and time is unknown, set this parameter to zero.
' CacheEntryType    [in] Unsigned long integer value that contains the cache type bitmask. This can be a combination of the following values: COOKIE_CACHE_ENTRY, EDITED_CACHE_ENTRY, NORMAL_CACHE_ENTRY, SPARCE_CACHE_ENTRY, STICKY_CACHE_ENTRY, TRACK_CACHE_ENTRY, TRACK_OFFLINE_CACHE_ENTRY, TRACK_ONLINE_CACHE_ENTRY, URLHISTORY_CACHE_ENTRY
' lpHeaderInfo      [in] Address of the buffer containing the header information. If this parameter is not NULL, the header information is treated as extended attributes of the URL that are returned in the INTERNET_CACHE_ENTRY_INFO structure.
' dwHeaderSize      [in] Unsigned long integer value that contains the size of the header information in TCHAR. If lpHeaderInfo is not NULL, this value is assumed to indicate the size of the buffer that will store the header information. An application can maintain headers as part of the data and provide dwHeaderSize together with a NULL value for lpHeaderInfo.
' lpszFileExtension [in] Reserved. Must be set to NULL.
' lpszOriginalUrl   [in] Address of a string variable that contains the original URL if redirection has occurred.
'
' Return:
' ¯¯¯¯¯¯¯
' Returns TRUE if successful, or FALSE otherwise. To get extended error information, call GetLastError . Possible error values include:
'   ERROR_DISK_FULL       The cache storage is full.
'   ERROR_FILE_NOT_FOUND  The specified local file is not found.
' ____________________________________________________________________________________________________________
' BOOL CommitUrlCacheEntry (LPCTSTR lpszUrlName, LPCTSTR lpszLocalFileName, FILETIME ExpireTime, FILETIME LastModifiedTime, DWORD CacheEntryType, LPCTSTR lpHeaderInfo, DWORD dwHeaderSize, LPCTSTR lpszFileExtension, LPCTSTR lpszOriginalUrl);
'=============================================================================================================
Public Declare Function CommitUrlCacheEntry Lib "wininet.dll" Alias "CommitUrlCacheEntryA" (ByVal lpszUrlName As String, ByVal lpszLocalFileName As String, ByRef ExpireTime As Currency, ByRef LastModifiedTime As Currency, ByVal CacheEntryType As Long, ByVal lpHeaderInfo As String, ByVal dwHeaderSize As Long, ByVal lpszFileExtension As String, ByVal lpszOriginalUrl As String) As Long 'BOOL


'=============================================================================================================
' CreateUrlCacheEntry
'
' Minimum Availability : Internet Explorer 3.0
'
' Parameters:
' ¯¯¯¯¯¯¯¯¯¯¯
' lpszUrlName        [in] Address of a string value that contains the name of the URL. The string should not contain any escape characters.
' dwExpectedFileSize [in] Unsigned long integer value that contains the expected size of the file needed to store the data corresponding to the source entity in TCHAR. If the expected size is unknown, set this value to zero.
' lpszFileExtension  [in] Address of a string value that contains an extension name of the file in the local storage.
' lpszFileName       [out] Address of a buffer that receives the file name. The buffer should be large enough (MAX_PATH) to store the path of the created file.
' dwReserved         [in] Reserved. Must be set to zero.
'
' Return:
' ¯¯¯¯¯¯¯
' Returns TRUE if successful, or FALSE otherwise. To get extended error information, call GetLastError.
' ____________________________________________________________________________________________________________
' BOOL CreateUrlCacheEntry (LPCTSTR lpszUrlName, DWORD dwExpectedFileSize, LPCTSTR lpszFileExtension, LPTSTR lpszFileName, DWORD dwReserved);
'=============================================================================================================
Public Declare Function CreateUrlCacheEntry Lib "wininet.dll" Alias "CreateUrlCacheEntryA" (ByVal lpszUrlName As String, ByVal dwExpectedFileSize As Long, ByVal lpszFileExtension As Long, ByVal lpszFileName As String, ByVal dwReserved As Long) As Long 'BOOL


'=============================================================================================================
' CreateUrlCacheGroup
'
' Minimum Availability : Internet Explorer 4.0
'
' Parameters:
' ¯¯¯¯¯¯¯¯¯¯¯
' dwFlags    [in] Unsigned long integer value that contains the flags to control the creation of the cache group. This can be set to CACHEGROUP_FLAG_GIDONLY, which causes CreateUrlCacheGroup to generate a unique GROUPID, but does not create a physical group.
' lpReserved [in] Reserved. Must be set to NULL.
'
' Return:
' ¯¯¯¯¯¯¯
' Returns a valid GROUPID if successful, or FALSE otherwise. To get specific error information, call GetLastError.
' ____________________________________________________________________________________________________________
' GROUPID CreateUrlCacheGroup (DWORD dwFlags, LPVOID lpReserved);
'=============================================================================================================
Public Declare Function CreateUrlCacheGroup Lib "wininet.dll" (ByVal dwFlags As Long, ByVal lpReserved As Long) As Currency


'=============================================================================================================
' DeleteUrlCacheEntry
'
' Minimum Availability : Internet Explorer 3.0
'
' Parameters:
' ¯¯¯¯¯¯¯¯¯¯¯
' lpszUrlName [in] Address of a string that contains the name of the source corresponding to the cache entry.
'
' Return:
' ¯¯¯¯¯¯¯
' Returns TRUE if successful, or FALSE otherwise. To get extended error information, call GetLastError . Possible error values include:
'   ERROR_ACCESS_DENIED   The file is locked or in use. The entry will be marked and will be deleted when the file is unlocked.
'   ERROR_FILE_NOT_FOUND  The file is not in the cache.
' ____________________________________________________________________________________________________________
' BOOL DeleteUrlCacheEntry (LPCTSTR lpszUrlName);
'=============================================================================================================
Public Declare Function DeleteUrlCacheEntry Lib "wininet.dll" Alias "DeleteUrlCacheEntryA" (ByVal lpszUrlName As String) As Long 'BOOL


'=============================================================================================================
' DeleteUrlCacheGroup
'
' Minimum Availability : Internet Explorer 4.0
'
' Parameters:
' ¯¯¯¯¯¯¯¯¯¯¯
' GroupId    [in] GROUPID value that is associated with the cache group to be released.
' dwFlags    [in] Unsigned long integer value containing the flags to control the cache group deletion. This can be set to CACHEGROUP_FLAG_FLUSHURL_ONDELETE, which causes DeleteUrlCacheGroup to delete all of the cache entries associated with this group, unless the entry belongs to another group.
' lpReserved [in] Reserved. Must be set to NULL.
'
' Return:
' ¯¯¯¯¯¯¯
' Returns TRUE if successful, or FALSE otherwise. To get specific error information, call GetLastError.
' ____________________________________________________________________________________________________________
' BOOL DeleteUrlCacheGroup (GROUPID GroupId, DWORD dwFlags, LPVOID lpReserved);
'=============================================================================================================
Public Declare Function DeleteUrlCacheGroup Lib "wininet.dll" (ByVal GroupId As Currency, ByVal dwFlags As Long, ByVal lpReserved As Long) As Long 'BOOL


'=============================================================================================================
' FindCloseUrlCache
'
' Minimum Availability : Internet Explorer 3.0
'
' Parameters:
' ¯¯¯¯¯¯¯¯¯¯¯
' hEnumHandle [in] Handle returned by a previous call to the FindFirstUrlCacheEntry function.
'
' Return:
' ¯¯¯¯¯¯¯
' Returns TRUE if successful, or FALSE otherwise. To get extended error information, call GetLastError.
' ____________________________________________________________________________________________________________
' BOOL FindCloseUrlCache (Handle hEnumHandle);
'=============================================================================================================
Public Declare Function FindCloseUrlCache Lib "wininet.dll" (ByVal hEnumHandle As Long) As Long 'BOOL


'=============================================================================================================
' FindFirstUrlCacheEntry
'
' Minimum Availability : Internet Explorer 3.0
'
' Parameters:
' ¯¯¯¯¯¯¯¯¯¯¯
' lpszUrlSearchPattern              [in]      Address of a string that contains the source name pattern to search for. This can be set to "cookie:" or "visited:" to enumerate the cookies and URL History entries in the cache. If this parameter is NULL, the function uses *.*.
' lpFirstCacheEntryInfo             [out]     Address of an INTERNET_CACHE_ENTRY_INFO structure.
' lpdwFirstCacheEntryInfoBufferSize [in, out] Address of an unsigned long integer variable that specifies the size of the lpFirstCacheEntryInfo buffer, in TCHARs. When the function returns, the variable contains the number of TCHARs copied to the buffer, or the required size, in bytes, needed to retrieve the cache entry.
'
' Return:
' ¯¯¯¯¯¯¯
' Returns a handle that the application can use in the FindNextUrlCacheEntry function to retrieve subsequent entries in the cache. If the function fails, the return value is NULL. To get extended error information, call GetLastError.
' ERROR_INSUFFICIENT_BUFFER indicates that the size of lpFirstCacheEntryInfo as specified by lpdwFirstCacheEntryInfoBufferSize is not sufficient to contain all the information. The value returned in lpdwFirstCacheEntryInfoBufferSize indicates the buffer size necessary to contain all the information.
' ____________________________________________________________________________________________________________
' HANDLE FindFirstUrlCacheEntry (LPCTSTR lpszUrlSearchPattern, LPINTERNET_CACHE_ENTRY_INFO lpFirstCacheEntryInfo, LPDWORD lpdwFirstCacheEntryInfoBufferSize);
'=============================================================================================================
Public Declare Function FindFirstUrlCacheEntry Lib "wininet.dll" Alias "FindFirstUrlCacheEntryA" (ByRef lpszUrlSearchPattern As String, ByRef lpFirstCacheEntryInfo As INTERNET_CACHE_ENTRY_INFO, ByVal lpdwFirstCacheEntryInfoBufferSize As Long) As Long


'=============================================================================================================
' FindFirstUrlCacheEntryEx
'
' Minimum Availability : Internet Explorer 4.0
'
' Parameters:
' ¯¯¯¯¯¯¯¯¯¯¯
' lpszUrlSearchPattern              [in]      BSTRthat contains the search pattern. Search patterns are currently not supported, so the value must be set to NULL to indicate all entries with the matching GROUPID.
' dwFlags                           [in]      Unsigned long integer value that contains the flags controlling the enumeration. No flags are currently implemented; this must be set to zero.
' dwFilter                          [in]      Unsigned long integer value that indicates the cache entry types that are allowed. This can be any combination of cache entry types: COOKIE_CACHE_ENTRY, NORMAL_CACHE_ENTRY, STICKY_CACHE_ENTRY, TRACK_OFFLINE_CACHE_ENTRY, TRACK_ONLINE_CACHE_ENTRY, URLHISTORY_CACHE_ENTRY,
' GroupId                           [in]      GROUPID value that indicates the cache group to enumerate. Set the value to zero to enumerate all entries that are not grouped.
' lpFirstCacheEntryInfo             [out]     Address of the buffer to hold the INTERNET_CACHE_ENTRY_INFO structure in which the cache entry information will be stored.
' lpdwFirstCacheEntryInfoBufferSize [in, out] Address of an unsigned long integer variable that indicates the size of lpFirstCacheEntryInfo, in TCHARs.
' lpReserved                        [out]     Reserved. Must be set to NULL.
' pcbReserved2                      [in, out] Reserved. Must be set to NULL.
' lpReserved3                       [in]      Reserved. Must be set to NULL.
'
' Return:
' ¯¯¯¯¯¯¯
' Returns a valid handle if successful, or NULL otherwise. To get specific error information, call GetLastError. If the function finds no matching files, GetLastError returns ERROR_NO_MORE_FILES.
' ____________________________________________________________________________________________________________
' HANDLE FindFirstUrlCacheEntryEx (LPCSTR lpszUrlSearchPattern, DWORD dwFlags, DWORD dwFilter, GROUPID GroupId, LPINTERNET_CACHE_ENTRY_INFO lpFirstCacheEntryInfo, LPDWORD lpdwFirstCacheEntryInfoBufferSize, LPVOID lpReserved, LPDWORD pcbReserved2, LPVOID lpReserved3);
'=============================================================================================================
Public Declare Function FindFirstUrlCacheEntryEx Lib "wininet.dll" Alias "FindFirstUrlCacheEntryExA" (ByVal lpszUrlSearchPattern As String, ByVal dwFlags As Long, ByVal dwFilter As Long, ByVal GroupId As Currency, ByRef lpFirstCacheEntryInfo As INTERNET_CACHE_ENTRY_INFO, ByRef lpdwFirstCacheEntryInfoBufferSize As Long, ByRef lpReserved1 As Long, ByRef pcbReserved2 As Long, ByVal lpReserved3 As Long) As Long


'=============================================================================================================
' FindFirstUrlCacheGroup
'
' Minimum Availability : Internet Explorer 5
'
' Parameters:
' ¯¯¯¯¯¯¯¯¯¯¯
' dwFlags           [in]      Reserved. Must be set to zero.
' dwFilter          [in]      Unsigned long integer value that indicates what filters to use. This can be one of the following values: CACHEGROUP_SEARCH_ALL, CACHEGROUP_SEARCH_BYURL
' lpSearchCondition [in]      Reserved. Must be set to NULL.
' dwSearchCondition [in]      Reserved. Must be set to zero.
' lpGroupId         [out]     Address of a GROUPID variable that contains the identification of the first cache group that matches the search criteria.
' lpReserved        [in, out] Reserved. Must be set to NULL.
'
' Return:
' ¯¯¯¯¯¯¯
' Returns a valid handle if successful, or NULL otherwise. To get specific error information, call GetLastError . If the function finds no matching files, GetLastError returns ERROR_NO_MORE_FILES.
' ____________________________________________________________________________________________________________
' HANDLE FindFirstUrlCacheGroup (DWORD dwFlags, DWORD dwFilter, LPVOID lpSearchCondition, DWORD dwSearchCondition, GROUPID *lpGroupId, LPVOID lpReserved);
'=============================================================================================================
Public Declare Function FindFirstUrlCacheGroup Lib "wininet.dll" (ByVal Reserved1 As Long, ByVal dwFilter As Long, ByVal lpReserved2 As Long, ByVal lpReserved3 As Long, ByRef lpGroupId As Currency, ByRef lpReserved4 As Long) As Long


'=============================================================================================================
' FindNextUrlCacheEntry
'
' Minimum Availability : Internet Explorer 3.0
'
' Parameters:
' ¯¯¯¯¯¯¯¯¯¯¯
' hEnumHandle                      [in]      Enumeration handle obtained from a previous call to FindFirstUrlCacheEntry.
' lpNextCacheEntryInfo             [out]     Address of an INTERNET_CACHE_ENTRY_INFO structure that receives information about the cache entry.
' lpdwNextCacheEntryInfoBufferSize [in, out] Address of an unsigned long integer variable that specifies the size of the lpNextCacheEntryInfo buffer, in TCHARs. When the function returns, the variable contains the number of TCHARs copied to the buffer, or the size of the buffer (in bytes) required to retrieve the cache entry.
'
' Return:
' ¯¯¯¯¯¯¯
' Returns TRUE if successful, or FALSE otherwise. To get extended error information, call GetLastError . Possible error values include:
'   ERROR_INSUFFICIENT_BUFFER  The size of lpNextCacheEntryInfo as specified by lpdwNextCacheEntryInfoBufferSize is not sufficient to contain all the information. The value returned in lpdwNextCacheEntryInfoBufferSize indicates the buffer size necessary to contain all the information.
'   ERROR_NO_MORE_ITEMS        The enumeration completed.
' ____________________________________________________________________________________________________________
' BOOL FindNextUrlCacheEntry (HANDLE hEnumHandle, LPINTERNET_CACHE_ENTRY_INFO lpNextCacheEntryInfo, LPWORD lpdwNextCacheEntryInfoBufferSize);
'=============================================================================================================
Public Declare Function FindNextUrlCacheEntry Lib "wininet.dll" Alias "FindNextUrlCacheEntryA" (ByVal hEnumHandle As Long, ByRef lpNextCacheEntryInfo As INTERNET_CACHE_ENTRY_INFO, ByRef lpdwNextCacheEntryInfoBufferSize As Long) As Long 'BOOL


'=============================================================================================================
' FindNextUrlCacheEntryEx
'
' Minimum Availability : Internet Explorer 4.0
'
' Parameters:
' ¯¯¯¯¯¯¯¯¯¯¯
' hEnumHandle                      [in]      Handle returned by FindFirstUrlCacheEntryEx, which started a cache enumeration.
' lpFirstCacheEntryInfo             [out]     Address of the buffer to hold the INTERNET_CACHE_ENTRY_INFO structure in which the cache entry information will be stored.
' lpdwFirstCacheEntryInfoBufferSize [in, out] Address of an unsigned long integer value that indicates the size of the buffer in TCHAR.
' lpReserved                        [out]     Reserved. Must be set to NULL.
' pcbReserved2                      [in, out] Reserved. Must be set to NULL.
' lpReserved3                       [in]      Reserved. Must be set to NULL.
'
' Return:
' ¯¯¯¯¯¯¯
'Returns TRUE if successful, or FALSE otherwise. To get specific error information, call GetLastError.
' ____________________________________________________________________________________________________________
' BOOL FindNextUrlCacheEntryEx (HANDLE hEnumHandle, LPINTERNET_CACHE_ENTRY_INFO lpFirstCacheEntryInfo, LPDWORD lpdwFirstCacheEntryInfoBufferSize, LPVOID lpReserved, LPDWORD pcbReserved2, LPVOID lpReserved3);
'=============================================================================================================
Public Declare Function FindNextUrlCacheEntryEx Lib "wininet.dll" Alias "FindNextUrlCacheEntryExA" (ByVal hEnumHandle As Long, ByRef lpFirstCacheEntryInfo As INTERNET_CACHE_ENTRY_INFO, ByRef lpdwFirstCacheEntryInfoBufferSize As Long, ByRef lpReserved1 As Long, ByRef pcbReserved2 As Long, ByVal lpReserved3 As Long) As Long 'BOOL


'=============================================================================================================
' FindNextUrlCacheGroup
'
' Minimum Availability : Internet Explorer 5
'
' Parameters:
' ¯¯¯¯¯¯¯¯¯¯¯
' hFind      [in]      Valid cache group enumeration handle returned by FindFirstUrlCacheGroup.
' lpGroupId  [out]     Address of a GROUPID variable that contains the cache group identification.
' lpReserved [in, out] Reserved. Must be set to NULL.
'
' Return:
' ¯¯¯¯¯¯¯
' Returns TRUE if successful, or FALSE otherwise. To get specific error information, call GetLastError.
' ____________________________________________________________________________________________________________
' BOOL FindNextUrlCacheGroup (HANDLE hFind, GROUPID *lpGroupId, LPVOID lpReserved);
'=============================================================================================================
Public Declare Function FindNextUrlCacheGroup Lib "wininet.dll" (ByVal hFind As Long, ByRef lpGroupId As Currency, ByRef lpReserved As Long) As Long 'BOOL


'=============================================================================================================
' FtpCommand
'
' Minimum Availability : Internet Explorer 5
'
' Parameters:
' ¯¯¯¯¯¯¯¯¯¯¯
' hConnect        [in]  HINTERNET handle returned from a call to InternetConnect.
' fExpectResponse [in]  BOOLvalue that indicates whether or not the application expects a response from the FTP server. This must be set to TRUE if a response is expected, or FALSE otherwise.
' dwFlags         [in]  Unsigned long integer value that contains the flags that control this function. This can be set to one of the following values: FTP_TRANSFER_TYPE_ASCII, FTP_TRANSFER_TYPE_BINARY
' lpszCommand     [in]  Address of a string value that contains the command to send to the FTP server.
' dwContext       [in]  Address of an unsigned long integer value that contains an application-defined value that is used to identify the application context in callbacks.
' phFtpCommand    [out] Address of an HINTERNET handle that will be created if a valid data socket is opened. The fExpectResponse parameter must be set to TRUE for phFtpCommand to be filled.
'
' Return:
' ¯¯¯¯¯¯¯
' Returns TRUE if successful, or FALSE otherwise. To get a specific error message, call GetLastError.
' ____________________________________________________________________________________________________________
' BOOL FtpCommand (HINTERNET hConnect, BOOL fExpectResponse, DWORD dwFlags, LPCTSTR lpszCommand, DWORD_PTR dwContext, HINTERNET *phFtpCommand);
'=============================================================================================================
Public Declare Function FtpCommand Lib "wininet.dll" Alias "FtpCommandA" (ByVal hConnect As Long, ByVal fExpectResponse As Long, ByVal dwFlags As Long, ByVal lpszCommand As String, ByRef dwContext As Long, ByRef phFtpCommand As Long) As Long 'BOOL


'=============================================================================================================
' FtpCreateDirectory
'
' Minimum Availability : Internet Explorer 3.0
'
' Parameters:
' ¯¯¯¯¯¯¯¯¯¯¯
' hConnect      [in] Valid HINTERNET handle returned by a previous call to InternetConnect using INTERNET_SERVICE_FTP.
' lpszDirectory [in] Address of a null-terminated string that contains the name of the directory to create on the remote system. This can be either a fully qualified path or a name relative to the current directory.
'
' Return:
' ¯¯¯¯¯¯¯
' Returns TRUE if successful, or FALSE otherwise. To get a specific error message, call GetLastError . If the error message indicates that the FTP server denied the request to create a directory, use InternetGetLastResponseInfo to determine why.
' ____________________________________________________________________________________________________________
' BOOL FtpCreateDirectory (HINTERNET hConnect, LPCTSTR lpszDirectory);
'=============================================================================================================
Public Declare Function FtpCreateDirectory Lib "wininet.dll" Alias "FtpCreateDirectoryA" (ByVal hConnect As Long, ByVal lpszDirectory As String) As Long 'BOOL


'=============================================================================================================
' FtpDeleteFile
'
' Minimum Availability : Internet Explorer 3.0
'
' Parameters:
' ¯¯¯¯¯¯¯¯¯¯¯
' hConnect     [in] Valid HINTERNET handle returned by a previous call to InternetConnect using INTERNET_SERVICE_FTP.
' lpszFileName [in] Address of a null-terminated string that contains the name of the file to delete on the remote system.
'
' Return:
' ¯¯¯¯¯¯¯
' Returns TRUE if successful, or FALSE otherwise. To get a specific error message, call GetLastError.
' ____________________________________________________________________________________________________________
' BOOL FtpDeleteFile (HINTERNET hConnect, LPCTSTR lpszFileName);
'=============================================================================================================
Public Declare Function FtpDeleteFile Lib "wininet.dll" Alias "FtpDeleteFileA" (ByVal hConnect As Long, ByVal lpszFileName As String) As Long 'BOOL


'=============================================================================================================
' FtpFindFirstFile
'
' Minimum Availability : Internet Explorer 3.0
'
' Parameters:
' ¯¯¯¯¯¯¯¯¯¯¯
' hConnect       [in]  Valid handle to an FTP session returned from InternetConnect.
' lpszSearchFile [in]  Address of a null-terminated string that specifies a valid directory path or file name for the FTP server's file system. The string can contain wildcards, but no blank spaces are allowed. If the value of lpszSearchFile is NULL or if it is an empty string, it will find the first file in the current directory on the server.
' lpFindFileData [out] Address of a WIN32_FIND_DATA structure that receives information about the found file or directory.
' dwFlags        [in]  Unsigned long integer value that contains the flags that control the behavior of this function. This can be a combination of the following values: INTERNET_FLAG_HYPERLINK, INTERNET_FLAG_NEED_FILE, INTERNET_FLAG_NO_CACHE_WRITE, INTERNET_FLAG_RELOAD, INTERNET_FLAG_RESYNCHRONIZE
' dwContext      [in]  Address of an unsigned long integer value that contains the application-defined value that associates this search with any application data. This parameter is used only if the application has already called InternetSetStatusCallback to set up a status callback function.
'
' Return:
' ¯¯¯¯¯¯¯
' Returns a valid handle for the request if the directory enumeration was started successfully, or returns NULL otherwise. To retrieve a specific error message, call GetLastError. If the function finds no matching files, GetLastError returns ERROR_NO_MORE_FILES.
' ____________________________________________________________________________________________________________
' HINTERNET FtpFindFirstFile (HINTERNET hConnect, LPCTSTR lpszSearchFile, LP lpFindFileData, DWORD dwFlags, DWORD_PTR dwContext);
'=============================================================================================================
Public Declare Function FtpFindFirstFile Lib "wininet.dll" Alias "FtpFindFirstFileA" (ByVal hConnect As Long, ByVal lpszSearchFile As String, ByRef lpFindFileData As WIN32_FIND_DATA, ByVal dwFlags As Long, ByRef dwContext As Long) As Long


'=============================================================================================================
' FtpGetCurrentDirectory
'
' Minimum Availability : Internet Explorer 3.0
'
' Parameters:
' ¯¯¯¯¯¯¯¯¯¯¯
' hConnect             [in]      Valid handle to an FTP session.
' lpszCurrentDirectory [out]     Address of a buffer that receives the current directory string, which specifies the absolute path to the current directory. The string is null-terminated.
' lpdwCurrentDirectory [in, out] Address of a variable that specifies the length, in characters, of the buffer for the current directory string. The buffer length must include room for a terminating NULL character. Using a length of MAX_PATH is sufficient for all paths. When the function returns, the variable receives the number of characters copied into the buffer.
'
' Return:
' ¯¯¯¯¯¯¯
' Returns TRUE if successful, or FALSE otherwise. To get a specific error message, call GetLastError .
' ____________________________________________________________________________________________________________
' BOOL FtpGetCurrentDirectory (HINTERNET hConnect, LPTSTR lpszCurrentDirectory, LPDWORD lpdwCurrentDirectory);
'=============================================================================================================
Public Declare Function FtpGetCurrentDirectory Lib "wininet.dll" Alias "FtpGetCurrentDirectoryA" (ByVal hConnect As Long, ByVal lpszCurrentDirectory As String, ByRef lpdwBufferLen As Long) As Long 'BOOL


'=============================================================================================================
' FtpGetFile
'
' Minimum Availability : Internet Explorer 3.0
'
' Parameters:
' ¯¯¯¯¯¯¯¯¯¯¯
' hConnect             [in] Valid handle to an FTP session.
' lpszRemoteFile       [in] Address of a null-terminated string that contains the name of the file to retrieve from the remote system.
' lpszNewFile          [in] Address of a null-terminated string that contains the name of the file to create on the local system.
' fFailIfExists        [in] BOOL that indicates whether the function should proceed if a local file of the specified name already exists. If fFailIfExists is TRUE and the local file exists, FtpGetFile fails.
' dwFlagsAndAttributes [in] Unsigned long integer value that contains the file attributes for the new file. This can be any combination of the FILE_ATTRIBUTE_* flags used by the CreateFile  function. For more information on FILE_ATTRIBUTE_* attributes, see CreateFile in the Platform SDK.
' dwFlags              [in] Unsigned long integer value that contains the flags that control how the function will handle the file download. The first set of flag values indicates the conditions under which the transfer occurs. These transfer type flags can be used in combination with the second set of flags that control caching.  The application can select one of these transfer type values: FTP_TRANSFER_TYPE_ASCII, FTP_TRANSFER_TYPE_BINARY, FTP_TRANSFER_TYPE_UNKNOWN, INTERNET_FLAG_TRANSFER_ASCII, INTERNET_FLAG_TRANSFER_BINARY, INTERNET_FLAG_HYPERLINK, INTERNET_FLAG_NEED_FILE, INTERNET_FLAG_RELOAD, INTERNET_FLAG_RESYNCHRONIZE
' dwContext            [in] Address of an unsigned long integer value that contains the application-defined value that associates this search with any application data. This is used only if the application has already called InternetSetStatusCallback to set up a status callback function.
'
' Return:
' ¯¯¯¯¯¯¯
' Returns TRUE if successful, or FALSE otherwise. To get a specific error message, call GetLastError.
' ____________________________________________________________________________________________________________
' BOOL FtpGetFile (HINTERNET hConnect, LPCTSTR lpszRemoteFile, LPCTSTR lpszNewFile, BOOL fFailIfExists, DWORD dwFlagsAndAttributes, DWORD dwFlags, DWORD_PTR dwContext);
'=============================================================================================================
Public Declare Function FtpGetFile Lib "wininet.dll" Alias "FtpGetFileA" (ByVal hConnect As Long, ByVal lpszRemoteFile As String, ByVal lpszLocalFilePath As String, ByVal fFailIfExists As Long, ByVal dwFlagsAndAttributes As Long, ByVal dwFlags As Long, ByRef dwContext As Long) As Long 'BOOL


'=============================================================================================================
' FtpGetFileSize
'
' Minimum Availability : Internet Explorer 5
'
' Parameters:
' ¯¯¯¯¯¯¯¯¯¯¯
' hFile            [in]  HINTERNET handle returned from a call to FtpOpenFile.
' lpdwFileSizeHigh [out] Pointer to the high-order unsigned long integer of the file size of the requested FTP resource.
'
' Return:
' ¯¯¯¯¯¯¯
' Returns the low-order unsigned long integer of the file size of the requested FTP resource.
' ____________________________________________________________________________________________________________
' DWORD FtpGetFileSize (HINTERNET hFile, LPDWORD lpdwFileSizeHigh);
'=============================================================================================================
Public Declare Function FtpGetFileSize Lib "wininet.dll" (ByVal hFile As Long, ByRef lpdwFileSizeHigh As Long) As Long


'=============================================================================================================
' FtpOpenFile
'
' Minimum Availability : Internet Explorer 3.0
'
' Parameters:
' ¯¯¯¯¯¯¯¯¯¯¯
' hConnect     [in] Valid HINTERNET handle to an FTP session.
' lpszFileName [in] Address of a null-terminated string that contains the name of the file to access on the remote system.
' dwAccess     [in] Unsigned long integer value that determines how the file will be accessed. This can be GENERIC_READ or GENERIC_WRITE, but not both.
' dwFlags      [in] Unsigned long integer value that contains the conditions under which the transfers occur. The application should select one transfer type and any of the flags that indicate how the caching of the file will be controlled.  The transfer type can be one of the following values: FTP_TRANSFER_TYPE_ASCII, FTP_TRANSFER_TYPE_BINARY, FTP_TRANSFER_TYPE_UNKNOWN, INTERNET_FLAG_TRANSFER_ASCII, INTERNET_FLAG_TRANSFER_BINARY, INTERNET_FLAG_HYPERLINK, INTERNET_FLAG_NEED_FILE, INTERNET_FLAG_RELOAD, INTERNET_FLAG_RESYNCHRONIZE
' dwContext    [in] Address of an unsigned long integer value that contains the application-defined value that associates this search with any application data. This is only used if the application has already called InternetSetStatusCallback to set up a status callback function.
'
' Return:
' ¯¯¯¯¯¯¯
' Returns a handle if successful, or NULL otherwise. To retrieve a specific error message, call GetLastError.
' ____________________________________________________________________________________________________________
' HINTERNET FtpOpenFile (HINTERNET hConnect, LPCTSTR lpszFileName, DWORD dwAccess, DWORD dwFlags, DWORD_PTR dwContext);
'=============================================================================================================
Public Declare Function FtpOpenFile Lib "wininet.dll" Alias "FtpOpenFileA" (ByVal hConnect As Long, ByVal lpszFileName As String, ByVal dwAccess As Long, ByVal dwFlags As Long, ByRef dwContext As Long) As Long


'=============================================================================================================
' FtpPutFile
'
' Minimum Availability : Internet Explorer 3.0
'
' Parameters:
' ¯¯¯¯¯¯¯¯¯¯¯
' hConnect          [in] Valid HINTERNET handle to an FTP session.
' lpszLocalFile     [in] Address of a null-terminated string that contains the name of the file to send from the local system.
' lpszNewRemoteFile [in] Address of a null-terminated string that contains the name of the file to create on the remote system.
' dwFlags           [in] Unsigned long integer value that contains the conditions under which the transfers occur. The application should select one transfer type and any of the flags that control how the caching of the file will be controlled.  The transfer type can be any one of the following values: FTP_TRANSFER_TYPE_ASCII, FTP_TRANSFER_TYPE_BINARY, FTP_TRANSFER_TYPE_UNKNOWN, INTERNET_FLAG_TRANSFER_ASCII, INTERNET_FLAG_TRANSFER_BINARY, INTERNET_FLAG_HYPERLINK, INTERNET_FLAG_NEED_FILE, INTERNET_FLAG_RELOAD, INTERNET_FLAG_RESYNCHRONIZE
' dwContext         [in] Address of an unsigned long integer value that contains the application-defined value that associates this search with any application data. This parameter is used only if the application has already called InternetSetStatusCallback to set up a status callback.
'
' Return:
' ¯¯¯¯¯¯¯
' Returns TRUE if successful, or FALSE otherwise. To get a specific error message, call GetLastError.
' ____________________________________________________________________________________________________________
' BOOL FtpPutFile (HINTERNET hConnect, LPCTSTR lpszLocalFile, LPCTSTR lpszNewRemoteFile, DWORD dwFlags, DWORD_PTR dwContext);
'=============================================================================================================
Public Declare Function FtpPutFile Lib "wininet.dll" Alias "FtpPutFileA" (ByVal hConnect As Long, ByVal lpszLocalFilePath As String, ByVal lpszRemoteFile As String, ByVal dwFlags As Long, ByRef dwContext As Long) As Long 'BOOL


'=============================================================================================================
' FtpRemoveDirectory
'
' Minimum Availability : Internet Explorer 3.0
'
' Parameters:
' ¯¯¯¯¯¯¯¯¯¯¯
' hConnect      [in] Valid HINTERNET handle to an FTP session.
' lpszDirectory [in] Address of a null-terminated string that contains the name of the directory to remove on the remote system. This can be either a fully qualified path or a name relative to the current directory.
'
' Return:
' ¯¯¯¯¯¯¯
' Returns TRUE if successful, or FALSE otherwise. To get a specific error message, call GetLastError . If the error message indicates that the FTP server denied the request to remove a directory, use InternetGetLastResponseInfo to determine why.
' ____________________________________________________________________________________________________________
' BOOL FtpRemoveDirectory (HINTERNET hConnect, LPCTSTR lpszDirectory);
'=============================================================================================================
Public Declare Function FtpRemoveDirectory Lib "wininet.dll" Alias "FtpRemoveDirectoryA" (ByVal hConnect As Long, ByVal lpszDirectory As String) As Long   'BOOL


'=============================================================================================================
' FtpRenameFile
'
' Minimum Availability : Internet Explorer 3.0
'
' Parameters:
' ¯¯¯¯¯¯¯¯¯¯¯
' hConnect     [in] Valid HINTERNET handle to an FTP session.
' lpszExisting [in] Address of a null-terminated string that contains the name of the file that will have its name changed on the remote FTP server.
' lpszNew      [in] Address of a null-terminated string that contains the new name for the remote file.
'
' Return:
' ¯¯¯¯¯¯¯
' Returns TRUE if successful, or FALSE otherwise. To get a specific error message, call GetLastError .
' ____________________________________________________________________________________________________________
' BOOL FtpRenameFile (HINTERNET hConnect, LPCTSTR lpszExisting, LPCTSTR lpszNew);
'=============================================================================================================
Public Declare Function FtpRenameFile Lib "wininet.dll" Alias "FtpRenameFileA" (ByVal hConnect As Long, ByVal lpszCurrentName As String, ByVal lpszNewName As String) As Long 'BOOL


'=============================================================================================================
' FtpSetCurrentDirectory
'
' Minimum Availability : Internet Explorer 3.0
'
' Parameters:
' ¯¯¯¯¯¯¯¯¯¯¯
' hConnect      [in] Valid HINTERNET handle to an FTP session.
' lpszDirectory [in] Address of a null-terminated string that contains the name of the directory to change to on the remote system. This can be either a fully qualified path or a name relative to the current directory.
'
' Return:
' ¯¯¯¯¯¯¯
' Returns TRUE if successful, or FALSE otherwise. To get a specific error message, call GetLastError . If the error message indicates that the FTP server denied the request to change a directory, use InternetGetLastResponseInfo to determine why.
' ____________________________________________________________________________________________________________
' BOOL FtpSetCurrentDirectory (HINTERNET hConnect, LPCTSTR lpszDirectory);
'=============================================================================================================
Public Declare Function FtpSetCurrentDirectory Lib "wininet.dll" Alias "FtpSetCurrentDirectoryA" (ByVal hConnect As Long, ByVal lpszDirectory As String) As Long 'BOOL


'=============================================================================================================
' GetUrlCacheEntryInfo
'
' Minimum Availability : Internet Explorer 3.0
'
' Parameters:
' ¯¯¯¯¯¯¯¯¯¯¯
' lpszUrlName                  [in]      Address of a string that contains the name of the cache entry. The name string should not contain any escape characters.
' lpCacheEntryInfo             [in]      Address of an INTERNET_CACHE_ENTRY_INFO structure that receives information about the cache entry.
' lpdwCacheEntryInfoBufferSize [in, out] Address of an unsigned long integer variable that specifies the size of the lpCacheEntryInfo buffer, in TCHARs. When the function returns, the variable contains the number of TCHARs copied to the buffer, or the required size of the buffer, in bytes.
'
' Return:
' ¯¯¯¯¯¯¯
' Returns TRUE if successful, or FALSE otherwise. To get extended error information, call GetLastError . Possible error values include:
'   ERROR_FILE_NOT_FOUND The specified cache entry is not found in the cache.
'   ERROR_INSUFFICIENT_BUFFER The size of lpCacheEntryInfo as specified by lpdwCacheEntryInfoBufferSize is not sufficient to contain all the information. The value returned in lpdwCacheEntryInfoBufferSize indicates the buffer size necessary to contain all the information.
' ____________________________________________________________________________________________________________
' BOOL GetUrlCacheEntryInfo (LPCTSTR lpszUrlName, LPINTERNET_CACHE_ENTRY_INFO lpCacheEntryInfo, LPDWORD lpdwCacheEntryInfoBufferSize);
'=============================================================================================================
Public Declare Function GetUrlCacheEntryInfo Lib "wininet.dll" Alias "GetUrlCacheEntryInfoA" (ByVal lpszUrlName As String, ByRef lpCacheEntryInfo As INTERNET_CACHE_ENTRY_INFO, ByRef lpdwCacheEntryInfoBufferSize As Long) As Long 'BOOL


'=============================================================================================================
' GetUrlCacheEntryInfoEx
'
' Minimum Availability : Internet Explorer 4.0
'
' Parameters:
' ¯¯¯¯¯¯¯¯¯¯¯
' lpszUrl                   [in] Address of a string that contains the name of the cache entry. The name string should not contain any escape characters.
' lpCacheEntryInfo          [out] Address of an INTERNET_CACHE_ENTRY_INFO structure that receives information about the cache entry.
' lpdwCacheEntryInfoBufSize [in, out] Address of an unsigned long integer variable that specifies the size of the lpCacheEntryInfo buffer, in TCHARs. When the function returns, the variable contains the number of TCHARs copied to the buffer, or the required size of the buffer in bytes.
' lpszReserved              [out] Reserved. Must be set to NULL.
' lpdwReserved              [in, out] Reserved. Must be set to NULL.
' lpReserved                Reserved. Must be set to NULL.
' dwFlags                   Reserved. Must be set to zero.
'
' Return:
' ¯¯¯¯¯¯¯
' Returns TRUE if the URL was located, or FALSE otherwise. Call GetLastError  for specific error information. Possible errors include:
'   ERROR_FILE_NOT_FOUND The URL was not found in the cache index, even after taking any cached redirections into account.
'   ERROR_INSUFFICIENT_BUFFER The buffer referenced by lpCacheEntryInfo was not large enough to hold the requested information. The size of the buffer needed will be returned to lpdwCacheEntryInfoBufSize.
' ____________________________________________________________________________________________________________
' BOOL GetUrlCacheEntryInfoEx (LPCTSTR lpszUrl, LPINTERNET_CACHE_ENTRY_INFO lpCacheEntryInfo, LPDWORD lpdwCacheEntryInfoBufSize, LPTSTR lpszReserved, LPDWORD lpdwReserved, LPVOID lpReserved, DWORD dwFlags);
'=============================================================================================================
Public Declare Function GetUrlCacheEntryInfoEx Lib "wininet.dll" Alias "GetUrlCacheEntryInfoExA" (ByVal lpszUrl As String, ByRef lpCacheEntryInfo As INTERNET_CACHE_ENTRY_INFO, ByRef lpdwCacheEntryInfoBufSize As Long, ByRef lpszReserved As String, ByRef lpdwReserved As Long, ByVal lpReserved As Long, ByVal dwFlags As Long) As Long


'=============================================================================================================
' GetUrlCacheGroupAttribute
'
' Minimum Availability : Internet Explorer 5
'
' Parameters:
' ¯¯¯¯¯¯¯¯¯¯¯
' gID           [in]      GROUPID of the cache group.
' dwFlags       [in]      Reserved. Must be set to zero.
' dwAttributes  [in]      Unsigned long integer value that contains the attributes to retrieve. This can be one of the following values: CACHEGROUP_ATTRIBUTE_BASIC, CACHEGROUP_ATTRIBUTE_FLAG, CACHEGROUP_ATTRIBUTE_GET_ALL, CACHEGROUP_ATTRIBUTE_GROUPNAME, CACHEGROUP_ATTRIBUTE_QUOTA, CACHEGROUP_ATTRIBUTE_STORAGE, CACHEGROUP_ATTRIBUTE_TYPE
' lpGroupInfo   [out]     Address of buffer that contains an INTERNET_CACHE_GROUP_INFO structure to store the requested information.
' lpdwGroupInfo [in, out] Address of an unsigned long integer value that contains the size of the lpGroupInfo buffer.
' lpReserved    [in, out] Reserved. Must be set to NULL.
'
' Return:
' ¯¯¯¯¯¯¯
' Returns TRUE if successful, or FALSE otherwise. To get specific error information, call GetLastError.
' ____________________________________________________________________________________________________________
' BOOL GetUrlCacheGroupAttribute (GROUPID gID, DWORD dwFlags, DWORD dwAttributes, LPINTERNET_CACHE_GROUP_INFO lpGroupInfo, LPDWORD lpdwGroupInfo, LPVOID lpReserved);
'=============================================================================================================
Public Declare Function GetUrlCacheGroupAttribute Lib "wininet.dll" Alias "GetUrlCacheGroupAttributeA" (ByVal gID As Currency, ByVal dwFlags As Long, ByVal dwAttributes As Long, ByRef lpGroupInfo As INTERNET_CACHE_ENTRY_INFO, ByRef lpdwGroupInfo As Long, ByRef lpReserved As Long) As Long


'=============================================================================================================
' GopherCreateLocator
'
' Minimum Availability : Internet Explorer 3.0
'
' Parameters:
' ¯¯¯¯¯¯¯¯¯¯¯
' lpszHost           [in]      Address of a string that contains the name of the host, or a dotted-decimal IP address (such as 198.105.232.1).
' nServerPort        [in]      Port number on which the Gopher server at lpszHost lives, in host byte order. If nServerPort is INTERNET_INVALID_PORT_NUMBER, the default Gopher port is used.
' lpszDisplayString  [in]      Address of a string that contains the Gopher document or directory to be displayed. If this parameter is NULL, the function returns the default directory for the Gopher server.
' lpszSelectorString [in]      Address of the selector string to send to the Gopher server in order to retrieve information. This parameter can be NULL.
' dwGopherType       [in]      Unsigned long integer value that specifies whether lpszSelectorString refers to a directory or document, and whether the request is Gopher+ or Gopher. The default value, GOPHER_TYPE_DIRECTORY, is used if the value of dwGopherType is zero. This can be one of the Gopher Type Values.
' lpszLocator        [out]     Address of a buffer that receives the locator string. If lpszLocator is NULL, lpdwBufferLength receives the necessary buffer length, but the function performs no other processing.
' lpdwBufferLength   [in, out] Address of an unsigned long integer value that contains the length of the lpszLocator buffer, in TCHARs. When the function returns, this parameter receives the number of TCHARs written to the lpszLocator buffer. If GetLastError  returns ERROR_INSUFFICIENT_BUFFER, this parameter receives the number of bytes required to form the locator successfully.
'
' Return:
' ¯¯¯¯¯¯¯
' Returns TRUE if successful, or FALSE otherwise. To get extended error information, call GetLastError or InternetGetLastResponseInfo.
' ____________________________________________________________________________________________________________
' BOOL GopherCreateLocator (LPCSTR lpszHost, INTERNET_PORT nServerPort, LPCTSTR lpszDisplayString, LPCTSTR lpszSelectorString, DWORD dwGopherType, LPTSTR lpszLocator, LPDWORD lpdwBufferLength);
'=============================================================================================================
Public Declare Function GopherCreateLocator Lib "wininet.dll" Alias "GopherCreateLocatorA" (ByVal lpszHost As String, ByVal nServerPort As Integer, ByVal lpszDisplayString As String, ByVal lpszSelectorString As String, ByVal dwGopherType As Long, ByRef lpszLocator As String, ByRef lpdwBufferLength As Long) As Long


'=============================================================================================================
' GopherFindFirstFile
'
' Minimum Availability : Internet Explorer 3.0
'
' Parameters:
' ¯¯¯¯¯¯¯¯¯¯¯
' hConnect         [in]  Handle to a Gopher session returned by InternetConnect.
' lpszLocator      [in]  Address of a string that contains the name of the item to locate. This can be one of the following:
'                          - Gopher locator returned by a previous call to this function or the InternetFindNextFile function.
'                          - NULL pointer or zero-length string indicating that the topmost information from a Gopher server is being returned.
'                          - Locator created by the GopherCreateLocator function.
' lpszSearchString [in]  Address of a buffer that contains the strings to search, if this request is to an index server. Otherwise, this parameter should be NULL.
' lpFindData       [out] Address of a GOPHER_FIND_DATA structure that receives the information retrieved by this function.
' dwFlags          [in]  Unsigned long integer value that contains the flags controlling the function behavior. This can be a combination of the following values: INTERNET_FLAG_HYPERLINK, INTERNET_FLAG_NEED_FILE, INTERNET_FLAG_NO_CACHE_WRITE, INTERNET_FLAG_RELOAD, INTERNET_FLAG_RESYNCHRONIZE,
' dwContext        [in]  Address of an unsigned long integer value that contains the application-defined value that associates this search with any application data.
'
' Return:
' ¯¯¯¯¯¯¯
' Returns a valid search handle if successful, or NULL otherwise. To retrieve extended error information, call GetLastError  or InternetGetLastResponseInfo.
' ____________________________________________________________________________________________________________
' HINTERNET GopherFindFirstFile (HINTERNET hConnect, LPCTSTR lpszLocator, LPCTSTR lpszSearchString, LPGOPHER_FIND_DATA lpFindData, DWORD dwFlags, DWORD_PTR dwContext);
'=============================================================================================================
Public Declare Function GopherFindFirstFile Lib "wininet.dll" Alias "GopherFindFirstFileA" (ByVal hConnect As Long, ByVal lpszLocator As String, ByVal lpszSearchString As String, ByRef lpFindData As GOPHER_FIND_DATA, ByVal dwFlags As Long, ByRef dwContext As Long) As Long


'=============================================================================================================
' GopherGetAttribute
'
' Minimum Availability : Internet Explorer 3.0
'
' Parameters:
' ¯¯¯¯¯¯¯¯¯¯¯
' hConnect               [in]  Handle to a Gopher session returned by InternetConnect.
' lpszLocator            [in]  Address of a string that identifies the item at the Gopher server on which to return attribute information.
' lpszAttributeName      [in]  Address of a space-delimited string specifying the names of attributes to return. If lpszAttributeName is NULL, GopherGetAttribute returns information about all attributes.
' lpBuffer               [out] Address of an application-defined buffer from which attribute information is retrieved.
' dwBufferLength         [in]  Unsigned long integer value containing the size, in TCHAR, of the lpBuffer buffer.
' lpdwCharactersReturned [out] Address of an unsigned long integer value that contains the number of characters read into the lpBuffer buffer.
' lpfnEnumerator         [in]  Address of a callback function that enumerates each attribute of the locator. This parameter is optional. If it is NULL, all the Gopher attribute information is placed into lpBuffer. If lpfnEnumerator is specified, the callback function is called once for each attribute of the object.  The callback function receives the address of a single GOPHER_ATTRIBUTE_TYPE structure with each call. The enumeration callback function allows the application to avoid having to parse the Gopher attribute information.
' dwContext              [in]  Unsigned long integer value that contains the application-defined value that associates this operation with any application data.
'
' Return:
' ¯¯¯¯¯¯¯
' Returns TRUE if the request is satisfied, or FALSE otherwise. To get extended error information, call GetLastError  or InternetGetLastResponseInfo.
' ____________________________________________________________________________________________________________
' BOOL GopherGetAttribute (HINTERNET hConnect, LPCTSTR lpszLocator, LPCTSTR lpszAttributeName, LPBYTE lpBuffer, DWORD dwBufferLength, LPDWORD lpdwCharactersReturned, GOPHER_ATTRIBUTE_ENUMERATOR lpfnEnumerator, DWORD dwContext);
'=============================================================================================================
Public Declare Function GopherGetAttribute Lib "wininet.dll" Alias "GopherGetAttributeA" (ByVal hConnect As Long, ByVal lpszLocator As String, ByVal lpszAttributeName As String, ByRef lpBuffer As Any, ByVal dwBufferLength As Long, ByRef lpdwCharactersReturned As Long, ByVal lpfnEnumerator As Long, ByVal dwContext As Long) As Long 'BOOL


'=============================================================================================================
' GopherGetLocatorType
'
' Minimum Availability : Internet Explorer 3.0
'
' Parameters:
' ¯¯¯¯¯¯¯¯¯¯¯
' lpszLocator    [in] Address of the Gopher locator string to parse.
' lpdwGopherType [out] Address of an unsigned long integer variable that receives the type of the locator. The type is a bitmask that consists of a combination of the Gopher Type Values.
'
' Return:
' ¯¯¯¯¯¯¯
' Returns TRUE if successful, or FALSE otherwise. To get extended error information, call GetLastError.
' ____________________________________________________________________________________________________________
' BOOL GopherGetLocatorType (LPCTSTR lpszLocator, LPDWORD lpdwGopherType);
'=============================================================================================================
Public Declare Function GopherGetLocatorType Lib "wininet.dll" Alias "GopherGetLocatorTypeA" (ByVal lpszLocator As String, ByRef lpdwGopherType As Long) As Long 'BOOL


'=============================================================================================================
' GopherOpenFile
'
' Minimum Availability : Internet Explorer 3.0
'
' Parameters:
' ¯¯¯¯¯¯¯¯¯¯¯
' hConnect    [in] Handle to a Gopher session returned by InternetConnect.
' lpszLocator [in] Address of a string that identifies the file to open. Generally, this locator is returned from a call to GopherFindFirstFile or InternetFindNextFile. Because the Gopher protocol has no concept of a current directory, the locator is always fully qualified.
' lpszView    [in] Address of a string that describes the view to open if several views of the file exist on the server. If lpszView is NULL, the function uses the default file view.
' dwFlags     [in] Unsigned long integer value that contains the conditions under which subsequent transfers occur. This can be any of the following values: INTERNET_FLAG_HYPERLINK, INTERNET_FLAG_NEED_FILE, INTERNET_FLAG_NO_CACHE_WRITE, INTERNET_FLAG_RELOAD, INTERNET_FLAG_RESYNCHRONIZE
' dwContext   [in] Address of an unsigned long integer value that contains an application-defined value that associates this operation with any application data.
'
' Return:
' ¯¯¯¯¯¯¯
' Returns a handle if successful, or NULL if the file cannot be opened. To retrieve extended error information, call GetLastError  or InternetGetLastResponseInfo.
' ____________________________________________________________________________________________________________
' HINTERNET GopherOpenFile (HINTERNET hConnect, LPCTSTR lpszLocator, LPCTSTR lpszView, DWORD dwFlags, DWORD_PTR dwContext);
'=============================================================================================================
Public Declare Function GopherOpenFile Lib "wininet.dll" Alias "GopherOpenFileA" (ByVal hConnect As Long, ByVal lpszLocator As String, ByVal lpszView As String, ByVal dwFlags As Long, ByRef dwContext As Long) As Long


'=============================================================================================================
' HttpAddRequestHeaders
'
' Minimum Availability : Internet Explorer 3.0
'
' Parameters:
' ¯¯¯¯¯¯¯¯¯¯¯
' hConnect        [in] HINTERNET handle returned by a call to the HttpOpenRequest function.
' lpszHeaders     [in] Address of a string variable containing the headers to append to the request. Each header must be terminated by a CR/LF (carriage return/line feed) pair.
' dwHeadersLength [in] Unsigned long integer value that contains the length, in TCHAR, of lpszHeaders. If this parameter is -1L, the function assumes that lpszHeaders is zero-terminated (ASCIIZ), and the length is computed.
' dwModifiers     [in] Unsigned long integer value that contains the flags used to modify the semantics of this function. Can be a combination of the following values: HTTP_ADDREQ_FLAG_ADD, HTTP_ADDREQ_FLAG_ADD_IF_NEW, HTTP_ADDREQ_FLAG_COALESCE, HTTP_ADDREQ_FLAG_COALESCE_WITH_COMMA, HTTP_ADDREQ_FLAG_COALESCE_WITH_SEMICOLON, HTTP_ADDREQ_FLAG_REPLACE
'
' Return:
' ¯¯¯¯¯¯¯
' Returns TRUE if successful, or FALSE otherwise. To get extended error information, call GetLastError.
' ____________________________________________________________________________________________________________
' BOOL HttpAddRequestHeaders (HINTERNET hConnect, LPCTSTR lpszHeaders, DWORD dwHeadersLength, DWORD dwModifiers);
'=============================================================================================================
Public Declare Function HttpAddRequestHeaders Lib "wininet.dll" Alias "HttpAddRequestHeadersA" (ByVal hConnect As Long, ByVal lpszHeaders As String, ByVal dwHeadersLength As Long, ByVal dwModifiers As Long) As Long 'BOOL


'=============================================================================================================
' HttpEndRequest
'
' Minimum Availability : Internet Explorer 4.0
'
' Parameters:
' ¯¯¯¯¯¯¯¯¯¯¯
' hRequest     [in]  HINTERNET handle returned by HttpOpenRequest and sent by HttpSendRequestEx.
' lpBuffersOut [out] Reserved. Must be set to NULL.
' dwFlags      [in]  Unsigned long integer value that contains the flags that control this function. Can be one of the following values: HSR_ASYNC, HSR_SYNC, HSR_USE_CONTEXT, HSR_INITIATE, HSR_DOWNLOAD, HSR_CHUNKED
' dwContext    [in]  Unsigned long integer variable that contains the application-defined context value for applications that register a status callback function.
'
' Return:
' ¯¯¯¯¯¯¯
' Returns TRUE if successful, or FALSE otherwise. To get extended error information, call GetLastError.
' ____________________________________________________________________________________________________________
' BOOL HttpEndRequest (HINTERNET hRequest, LPINTERNET_BUFFERS lpBuffersOut, DWORD dwFlags, DWORD dwContext);
'=============================================================================================================
Public Declare Function HttpEndRequest Lib "wininet.dll" Alias "HttpEndRequestA" (ByVal hRequest As Long, ByRef lpBuffersOut As Long, ByVal dwFlags As Long, ByVal dwContext As Long) As Long


'=============================================================================================================
' HttpOpenRequest
'
' Minimum Availability : Internet Explorer 3.0
'
' Parameters:
' ¯¯¯¯¯¯¯¯¯¯¯
' hConnect        [in] HINTERNET handle to an HTTP session returned by InternetConnect.
' lpszVerb        [in] Address of a string that contains the verb to use in the request. If this parameter is NULL, the function uses GET as the verb.
' lpszObjectName  [in] Address of a string that contains the name of the target object of the specified verb. This is generally a file name, an executable module, or a search specifier.
' lpszVersion     [in] Address of a string that contains the HTTP version. If this parameter is NULL, the function uses HTTP/1.0 as the version.
' lpszReferer     [in] Address of a string that specifies the URL of the document from which the URL in the request (lpszObjectName) was obtained. If this parameter is NULL, no "referrer" is specified.
' lpszAcceptTypes [in] Address of a null-terminated array of string pointers indicating media types accepted by the client. If this parameter is NULL, no types are accepted by the client. Servers generally interpret a lack of accept types to indicate that the client accepts only documents of type "text/*" (that is, only text documents—no pictures or other binary files). For a list of valid media types, see "Media Types" at ftp://ftp.isi.edu/in-notes/iana/assignments/media-types/media-types .
' dwFlags         [in] Unsigned long integer value that contains the Internet flag values. This can be any of the following values: INTERNET_FLAG_CACHE_IF_NET_FAIL, INTERNET_FLAG_HYPERLINK, INTERNET_FLAG_IGNORE_CERT_CN_INVALID, INTERNET_FLAG_IGNORE_CERT_DATE_INVALID, INTERNET_FLAG_IGNORE_REDIRECT_TO_HTTP, INTERNET_FLAG_IGNORE_REDIRECT_TO_HTTPS, INTERNET_FLAG_KEEP_CONNECTION, INTERNET_FLAG_NEED_FILE, INTERNET_FLAG_NO_AUTH, INTERNET_FLAG_NO_AUTO_REDIRECT, INTERNET_FLAG_NO_CACHE_WRITE, INTERNET_FLAG_NO_COOKIES, INTERNET_FLAG_NO_UI, INTERNET_FLAG_PRAGMA_NOCACHE, INTERNET_FLAG_RELOAD, INTERNET_FLAG_RESYNCHRONIZE, INTERNET_FLAG_SECURE
' dwContext       [in] Address of an unsigned long integer value that contains the application-defined value that associates this operation with any application data.
'
' Return:
' ¯¯¯¯¯¯¯
' Returns a valid (non-NULL) HTTP request handle if successful, or NULL otherwise. To retrieve extended error information, call GetLastError.
' ____________________________________________________________________________________________________________
' HINTERNET HttpOpenRequest (HINTERNET hConnect, LPCTSTR lpszVerb, LPCTSTR lpszObjectName, LPCTSTR lpszVersion, LPCTSTR lpszReferer, LPCTSTR *lpszAcceptTypes, DWORD dwFlags, DWORD_PTR dwContext);
'=============================================================================================================
Public Declare Function HttpOpenRequest Lib "wininet.dll" Alias "HttpOpenRequestA" (ByVal hConnect As Long, ByVal lpszVerb As String, ByVal lpszObjectName As String, ByVal lpszVersion As String, ByVal lpszReferer As String, ByVal lpszAcceptTypes As String, ByVal dwFlags As Long, ByRef dwContext As Long) As Long


'=============================================================================================================
' HttpQueryInfo
'
' Minimum Availability : Internet Explorer 3.0
'
' Sample Use:
' ¯¯¯¯¯¯¯¯¯¯¯
' bRet = HttpQueryInfo(hResource, HTTP_QUERY_RAW_HEADERS_CRLF, lpvSomeBuffer, &dwSize, <dtype rid="NULL"/>));
'
' Parameters:
' ¯¯¯¯¯¯¯¯¯¯¯
' hRequest         [in]      HINTERNET request handle returned by HttpOpenRequest or InternetOpenUrl.
' dwInfoLevel      [in]      Unsigned long integer value that contains a combination of an attribute to retrieve and the flags that modify the request. The attribute can be any one of the Attributes, and the flag can be any one of the Modifiers on the Query Info Flags page.
' lpvBuffer        [in]      Address of the buffer that receives the information. This must not be set to NULL.
' lpdwBufferLength [in]      Address of a value that contains the length of the data buffer, in TCHARs. When the function returns, this parameter contains the address of a value specifying the length of the information written to the buffer. When the function returns strings, the following rules apply:
'   If the function succeeds, lpdwBufferLength specifies the length of the string, in TCHARs, minus 1 for the terminating NULL.
'   If the function fails and ERROR_INSUFFICIENT_BUFFER is returned, lpdwBufferLength specifies the number of bytes that the application must allocate to receive the string.
' lpdwIndex        [in, out] Address of a zero-based header index used to enumerate multiple headers with the same name. When calling the function, this parameter is the index of the specified header to return. When the function returns, this parameter is the index of the next header. If the next index cannot be found, ERROR_HTTP_HEADER_NOT_FOUND is returned.
'
' Return:
' ¯¯¯¯¯¯¯
' Returns TRUE if successful, or FALSE otherwise. To get extended error information, call GetLastError.
' ____________________________________________________________________________________________________________
' BOOL HttpQueryInfo (HINTERNET hRequest, DWORD dwInfoLevel, LPVOID lpvBuffer, LPDWORD lpdwBufferLength, LPDWORD lpdwIndex);
'=============================================================================================================
Public Declare Function HttpQueryInfo Lib "wininet.dll" Alias "HttpQueryInfoA" (ByVal hRequest As Long, ByVal dwInfoLevel As Long, ByVal lpvBuffer As Long, ByVal lpdwBufferLength As Long, ByRef lpdwIndex As Long) As Long  'BOOL


'=============================================================================================================
' HttpSendRequest
'
' Minimum Availability : Internet Explorer 3.0
'
' Parameters:
' ¯¯¯¯¯¯¯¯¯¯¯
' hRequest         [in] HINTERNET handle returned by HttpOpenRequest.
' lpszHeaders      [in] Address of a string variable that contains the additional headers to be appended to the request. This parameter can be NULL if there are no additional headers to append.
' dwHeadersLength  [in] Unsigned long integer value that contains the length, in TCHAR, of the additional headers. If this parameter is -1L and lpszHeaders is not NULL, the function assumes that lpszHeaders is zero-terminated (ASCIIZ), and the length is calculated.
' lpOptional       [in] Address of a buffer containing any optional data to send immediately after the request headers. This parameter is generally used for POST and PUT operations. The optional data can be the resource or information being posted to the server. This parameter can be NULL if there is no optional data to send.
' dwOptionalLength [in] Unsigned long integer value that contains the length, in bytes, of the optional data. This parameter can be zero if there is no optional data to send.
'
' Return:
' ¯¯¯¯¯¯¯
' Returns TRUE if successful, or FALSE otherwise. To get extended error information, call GetLastError .
' ____________________________________________________________________________________________________________
' BOOL HttpSendRequest (HINTERNET hRequest, LPCTSTR lpszHeaders, DWORD dwHeadersLength, LPVOID lpOptional, DWORD dwOptionalLength);
'=============================================================================================================
Public Declare Function HttpSendRequest Lib "wininet.dll" Alias "HttpSendRequestA" (ByVal hRequest As Long, ByVal lpszHeaders As String, ByVal dwHeadersLength As Long, ByVal lpOptional As Long, ByVal dwOptionalLength As Long) As Long 'BOOL


'=============================================================================================================
' HttpSendRequestEx
'
' Minimum Availability : Internet Explorer 4.0
'
' Parameters:
' ¯¯¯¯¯¯¯¯¯¯¯
' hRequest     [in]  HINTERNET handle returned by HttpOpenRequest.
' lpBuffersIn  [in]  Optional. Address of an INTERNET_BUFFERS structure.
' lpBuffersOut [out] Optional. Address of an INTERNET_BUFFERS structure.
' dwFlags      [in]  One of the following values: HSR_ASYNC, HSR_SYNC, HSR_USE_CONTEXT, HSR_INITIATE, HSR_DOWNLOAD, HSR_CHUNKED
' dwContext    [in]  Unsigned long integer variable that contains the application-defined context value, if a status callback function has been registered.
'
' Return:
' ¯¯¯¯¯¯¯
' Returns TRUE if successful, or FALSE otherwise. To get extended error information, call GetLastError .
' ____________________________________________________________________________________________________________
' BOOL HttpSendRequestEx (HINTERNET hRequest, LPINTERNET_BUFFERS lpBuffersIn, LPINTERNET_BUFFERS lpBuffersOut, DWORD dwFlags, DWORD dwContext);
'=============================================================================================================
Public Declare Function HttpSendRequestEx Lib "wininet.dll" Alias "HttpSendRequestExA" (ByVal hRequest As Long, ByRef lpBuffersIn As INTERNET_BUFFERS, ByRef lpBuffersOut As INTERNET_BUFFERS, ByVal dwFlags As Long, ByVal dwContext As Long) As Long


'=============================================================================================================
' InternetAttemptConnect
'
' Minimum Availability : Internet Explorer 3.0
'
' Parameters:
' ¯¯¯¯¯¯¯¯¯¯¯
' dwReserved [in] Reserved. Must be set to zero.
'
' Return:
' ¯¯¯¯¯¯¯
' Returns ERROR_SUCCESS if successful, or a Microsoft® Win32® error value otherwise.
' ____________________________________________________________________________________________________________
' DWORD InternetAttemptConnect (DWORD dwReserved);
'=============================================================================================================
Public Declare Function InternetAttemptConnect Lib "wininet.dll" (ByVal dwReserved As Long) As Long


'=============================================================================================================
' InternetAutodial
'
' Minimum Availability : Internet Explorer 4.0
'
' Parameters:
' ¯¯¯¯¯¯¯¯¯¯¯
'
' Return:
' ¯¯¯¯¯¯¯
' Returns TRUE if successful, or FALSE otherwise.
' ____________________________________________________________________________________________________________
' BOOL InternetAutodial (DWORD dwFlags, hWnd hwndParent);
'=============================================================================================================
Public Declare Function InternetAutodial Lib "wininet.dll" (ByVal dwFlags As Long, ByVal hWndParent As Long) As Long


'=============================================================================================================
' InternetAutodialHangup
'
' Minimum Availability : Internet Explorer 4.0
'
' Parameters:
' ¯¯¯¯¯¯¯¯¯¯¯
' dwReserved [in] Reserved. Must be set to zero.
'
' Return:
' ¯¯¯¯¯¯¯
' Returns TRUE if successful, or FALSE otherwise.
' ____________________________________________________________________________________________________________
' BOOL InternetAutodialHangup (DWORD dwReserved);
'=============================================================================================================
Public Declare Function InternetAutodialHangup Lib "wininet.dll" (ByVal dwReserved As Long) As Long


'=============================================================================================================
' InternetCanonicalizeUrl
'
' Minimum Availability : Internet Explorer 3.0
'
' Parameters:
' ¯¯¯¯¯¯¯¯¯¯¯
' lpszUrl          [in]      Address of the string that contains the URL to canonicalize.
' lpszBuffer       [out]     Address of the buffer that receives the resulting canonicalized URL.
' lpdwBufferLength [in, out] Address of an unsigned long integer value that contains the length, in TCHARs, of the lpszBuffer buffer. If the function succeeds, this parameter receives the length of the lpszBuffer buffer—the length does not include the terminating NULL character. If the function fails, this parameter receives the required length, in bytes, of the lpszBuffer buffer—the required length includes the terminating NULL character.
' dwFlags          [in]      Unsigned long integer value that contains the flags that control canonicalization. If no flags are specified (dwFlags = 0), the function converts all unsafe characters and meta sequences (such as \.,\ .., and \...) to escape sequences. dwFlags can be one of the following values: ICU_BROWSER_MODE, ICU_DECODE, ICU_ENCODE_PERCENT, ICU_ENCODE_SPACES_ONLY, ICU_NO_ENCODE, ICU_NO_META
'
' Return:
' ¯¯¯¯¯¯¯
' Returns TRUE if successful, or FALSE otherwise. To get extended error information, call GetLastError . Possible errors include:
'   ERROR_BAD_PATHNAME The URL could not be canonicalized. This flag is valid for Internet Explorer 5 and later versions of the Win32 Internet API.
'   ERROR_INSUFFICIENT_BUFFER The canonicalized URL is too large to fit in the buffer provided. The lpdwBufferLength parameter is set to the size, in bytes, of the buffer required to hold the canonicalized URL.
'   ERROR_INTERNET_INVALID_URL The format of the URL is invalid.
'   ERROR_INVALID_PARAMETER There is an invalid string, buffer, buffer size, or flags parameter.
' ____________________________________________________________________________________________________________
' BOOL InternetCanonicalizeUrl (LPCTSTR lpszUrl, LPTSTR lpszBuffer, LPDWORD lpdwBufferLength, DWORD dwFlags);
'=============================================================================================================
Public Declare Function InternetCanonicalizeUrl Lib "wininet.dll" Alias "InternetCanonicalizeUrlA" (ByVal lpszUrl As String, ByVal lpszBuffer As String, ByRef lpdwBufferLength As Long, ByVal dwFlags) As Long


'=============================================================================================================
' InternetCheckConnection
'
' Minimum Availability : Internet Explorer 3.0
'
' Parameters:
' ¯¯¯¯¯¯¯¯¯¯¯
' lpszUrl    [in] Address of a string containing the URL to use to check the connection. This value can be set to NULL.
' dwFlags    [in] Unsigned long integer value containing the flag values. FLAG_ICC_FORCE_CONNECTION is the only flag that is currently available. If this flag is set, it forces a connection. A sockets connection is attempted in the following order:
'   If lpszUrl is non-NULL, the host value is extracted from it and used to ping that specific host.
'   If lpszUrl is NULL and there is an entry in WinInet's internal server database for the nearest sever, the host value is extracted from the entry and used to ping that server.
' dwReserved [in] Reserved. Must be set to zero.
'
' Return:
' ¯¯¯¯¯¯¯
' Returns TRUE if a connection is made successfully, or FALSE otherwise. Use GetLastError  to retrieve the error code. ERROR_NOT_CONNECTED is returned by GetLastError if a connection cannot be made or if the sockets database is unconditionally offline.
' ____________________________________________________________________________________________________________
' BOOL InternetCheckConnection (LPCTSTR lpszUrl, DWORD dwFlags, DWORD dwReserved);
'=============================================================================================================
Public Declare Function InternetCheckConnection Lib "wininet.dll" Alias "InternetCheckConnectionA" (ByVal lpszUrl As String, ByVal dwFlags As Long, ByVal dwReserved As Long) As Long 'BOOL


'=============================================================================================================
' InternetCloseHandle
'
' Minimum Availability : Internet Explorer 3.0
'
' Parameters:
' ¯¯¯¯¯¯¯¯¯¯¯
' hInternet [in] Valid HINTERNET handle to be closed.
'
' Return:
' ¯¯¯¯¯¯¯
' Returns TRUE if the handle is successfully closed, or FALSE otherwise. To get extended error information, call GetLastError.
' ____________________________________________________________________________________________________________
' BOOL InternetCloseHandle (HINTERNET hInternet);
'=============================================================================================================
Public Declare Function InternetCloseHandle Lib "wininet.dll" (ByVal hInternet As Long) As Long 'BOOL


'=============================================================================================================
' InternetCombineUrl
'
' Minimum Availability : Internet Explorer 3.0
'
' Parameters:
' ¯¯¯¯¯¯¯¯¯¯¯
' lpszBaseUrl      [in]      Address of a string variable that contains the base URL.
' lpszRelativeUrl  [in]      Address of a string variable that contains the relative URL.
' lpszBuffer       [out]     Address of a buffer that receives the combined URL.
' lpdwBufferLength [in, out] Address of an unsigned long integer value that contains the size, in TCHARs, of the lpszBuffer buffer. If the function succeeds, this parameter receives the length, in TCHARs, of the combined URL—this length does not include the NULL terminator. If the function fails, this parameter receives the length, in bytes, of the required buffer—this length includes the NULL terminator.
' dwFlags          [in]      Unsigned long integer value that contains the flags controlling the operation of the function. This can be one of the following values: ICU_BROWSER_MODE, ICU_DECODE, ICU_ENCODE_PERCENT, ICU_ENCODE_SPACES_ONLY, ICU_NO_ENCODE, ICU_NO_META
'
' Return:
' ¯¯¯¯¯¯¯
' Returns TRUE if successful, or FALSE otherwise. To get extended error information, call GetLastError . Possible errors include:
'   ERROR_BAD_PATHNAME The URLs could not be combined.
'   ERROR_INSUFFICIENT_BUFFER The buffer supplied to the function was insufficient or NULL. The value indicated by the lpdwBufferLength parameter will contain the number of bytes required to hold the combined URL.
'   ERROR_INTERNET_INVALID_URL The format of the URL is invalid.
'   ERROR_INVALID_PARAMETER There is an invalid string, buffer, buffer size, or flags parameter.
' ____________________________________________________________________________________________________________
' BOOL InternetCombineUrl (LPCTSTR lpszBaseUrl, LPCTSTR lpszRelativeUrl, LPTSTR lpszBuffer, LPDWORD lpdwBufferLength, DWORD dwFlags);
'=============================================================================================================
Public Declare Function InternetCombineUrl Lib "wininet.dll" Alias "InternetCombineUrlA" (ByVal lpszBaseUrl As String, ByVal lpszRelativeUrl As Long, ByVal lpszBuffer As Long, ByVal lpdwBufferLength As Long, ByVal dwFlags As Long) As Long


'=============================================================================================================
' InternetConfirmZoneCrossing
'
' Minimum Availability : Internet Explorer 3.0
'
' Parameters:
' ¯¯¯¯¯¯¯¯¯¯¯
' hWnd      [in] Handle to the parent window for any needed dialog box.
' szUrlPrev [in] Address of a string variable containing the URL that was viewed before the current request was made.
' szUrlNew  [in] Address of a string variable containing the new URL that the user has requested to view.
' bPost     [in] BOOLvalue that determines if a post is being made by this request. If bPost is set to TRUE (1), a post is being made in this request. This flag is ignored in this release.
'
' Return:
' ¯¯¯¯¯¯¯
' Returns one of the following values:
'   ERROR_SUCCESS The user confirmed that it was okay to continue, or there was no user input needed.
'   ERROR_CANCELLED The user canceled the request.
'   ERROR_NOT_ENOUGH_MEMORY There is not enough memory to carry out the request.
' ____________________________________________________________________________________________________________
' DWORD InternetConfirmZoneCrossing (HWND hWnd, LPTSTR szUrlPrev, LPTSTR szUrlNew, BOOL bPost);
'=============================================================================================================
Public Declare Function InternetConfirmZoneCrossing Lib "wininet.dll" Alias "InternetConfirmZoneCrossingA" (ByVal hwnd As Long, ByVal szUrlPrev As String, ByVal szUrlNew As String, ByVal bPost As Long) As Long


'=============================================================================================================
' InternetConnect
'
' Minimum Availability : Internet Explorer 3.0
'
' Parameters:
' ¯¯¯¯¯¯¯¯¯¯¯
' hInternet      [in] Valid HINTERNET handle returned by a previous call to InternetOpen.
' lpszServerName [in] Address of a null-terminated string that contains the host name of an Internet server. Alternately, the string can contain the IP number of the site, in ASCII dotted-decimal format (for example, 11.0.1.45).
' nServerPort    [in] Long value representing the TCP/IP port on the server to connect to. These flags set only the port that will be used. The service is set by the value of dwService. This can be one of the following values: INTERNET_DEFAULT_FTP_PORT, INTERNET_DEFAULT_GOPHER_PORT, INTERNET_DEFAULT_HTTP_PORT, INTERNET_DEFAULT_HTTPS_PORT, INTERNET_DEFAULT_SOCKS_PORT, INTERNET_INVALID_PORT_NUMBER
' lpszUsername   [in] Address of a null-terminated string that contains the name of the user to log on. If this parameter is NULL, the function uses an appropriate default, except for HTTP; a NULL parameter in HTTP causes the server to return an error. For the FTP protocol, the default is "anonymous".
' lpszPassword   [in] Address of a null-terminated string that contains the password to use to log on. If both lpszPassword and lpszUsername are NULL, the function uses the default "anonymous" password. In the case of FTP, the default password is the user's e-mail name. If lpszPassword is NULL, but lpszUsername is not NULL, the function uses a blank password.
' dwService      [in] Unsigned long integer value that contains the type of service to access. This can be one of the following values: INTERNET_SERVICE_FTP, INTERNET_SERVICE_GOPHER, INTERNET_SERVICE_HTTP
' dwFlags        [in] Unsigned long integer value that contains the flags specific to the service used. When the value of dwService is INTERNET_SERVICE_FTP, INTERNET_FLAG_PASSIVE causes the application to use passive FTP semantics.
' dwContext      [in] Address of an unsigned long integer value that contains an application-defined value that is used to identify the application context for the returned handle in callbacks.
'
' Return:
' ¯¯¯¯¯¯¯
' Returns a valid handle to the FTP, Gopher, or HTTP session if the connection is successful, or NULL otherwise. To retrieve extended error information, call GetLastError . An application can also use InternetGetLastResponseInfo to determine why access to the service was denied.
' ____________________________________________________________________________________________________________
' HINTERNET InternetConnect (HINTERNET hInternet, LPCTSTR lpszServerName, INTERNET_PORT nServerPort, LPCTSTR lpszUsername, LPCTSTR lpszPassword, DWORD dwService, DWORD dwFlags, DWORD_PTR dwContext);
'=============================================================================================================
Public Declare Function InternetConnect Lib "wininet.dll" Alias "InternetConnectA" (ByVal hInternet As Long, ByVal lpszServerName As String, ByVal nServerPort As Integer, ByVal lpszUserName As String, ByVal lpszPassword As String, ByVal dwService As Long, ByVal dwFlags As Long, ByRef dwContext As Long) As Long


'=============================================================================================================
' InternetCrackUrl
'
' Minimum Availability : Internet Explorer 3.0
'
' Parameters:
' ¯¯¯¯¯¯¯¯¯¯¯
' lpszUrl         [in]      Address of a string that contains the canonical URL to crack.
' dwUrlLength     [in]      Unsigned long integer value that contains the length of the lpszUrl string in TCHAR, or zero if lpszUrl is an ASCIIZ string.
' dwFlags         [in]      Unsigned long integer value that contains the flags controlling the operation. This can be one of the following values: ICU_DECODE, ICU_ESCAPE
' lpUrlComponents [in, out] Address of a URL_COMPONENTS structure that receives the URL components.
'
' Return:
' ¯¯¯¯¯¯¯
' Returns TRUE if the function succeeds, or FALSE otherwise. To get extended error information, call GetLastError.
' ____________________________________________________________________________________________________________
' BOOL InternetCrackUrl (LPCTSTR lpszUrl, DWORD dwUrlLength, DWORD dwFlags, LPURL_COMPONENTS lpUrlComponents);
'=============================================================================================================
Public Declare Function InternetCrackUrl Lib "wininet.dll" Alias "InternetCrackUrlA" (ByVal lpszUrl As String, ByVal dwUrlLength As Long, ByVal dwFlags As Long, ByRef lpUrlComponents As URL_COMPONENTS) As Long


'=============================================================================================================
' InternetCreateUrl
'
' Minimum Availability : Internet Explorer 3.0
'
' Parameters:
' ¯¯¯¯¯¯¯¯¯¯¯
' lpUrlComponents [in]      Address of a URL_COMPONENTS structure that contains the components from which to create the URL.
' dwFlags         [in]      Unsigned long integer value that contains the flags that control the operation of this function. This can be one or both of these values: ICU_ESCAPE, ICU_USERNAME
' lpszUrl         [out]     Address of a buffer that receives the URL.
' lpdwUrlLength   [in, out] Address of an unsigned long integer value that contains the length, in TCHARs, of the lpszUrl buffer. When the function returns, this parameter receives the length, in TCHARs, of the URL string, minus 1 for the terminating character. If GetLastError  returns ERROR_INSUFFICIENT_BUFFER, this parameter receives the number of bytes required to hold the created URL.
'
' Return:
' ¯¯¯¯¯¯¯
' Returns TRUE if the function succeeds, or FALSE otherwise. To get extended error information, call GetLastError.
' ____________________________________________________________________________________________________________
' BOOL InternetCreateUrl (LPURL_COMPONENTS lpUrlComponents, DWORD dwFlags, LPTSTR lpszUrl, LPDWORD lpdwUrlLength);
'=============================================================================================================
Public Declare Function InternetCreateUrl Lib "wininet.dll" Alias "InternetCreateUrlA" (ByRef lpUrlComponents As URL_COMPONENTS, ByVal dwFlags As Long, ByVal lpszUrl As String, ByRef lpdwUrlLength As Long) As Long 'BOOL


'=============================================================================================================
' InternetDial
'
' Minimum Availability : Internet Explorer 4.0
'
' Parameters:
' ¯¯¯¯¯¯¯¯¯¯¯
' hwndParent     [in] Handle to the parent window.
' lpszConnectoid [in] Address of a string variable containing the name of the dial-up connection to use.
' dwFlags        [in] Unsigned long integer value that contains the flags to use. This can be one of the following values: INTERNET_AUTODIAL_FORCE_ONLINE, INTERNET_AUTODIAL_FORCE_UNATTENDED, INTERNET_DIAL_FORCE_PROMPT, INTERNET_DIAL_UNATTENDED, INTERNET_DIAL_SHOW_OFFLINE
' lpdwConnection [out] Address of an unsigned long integer value containing the number associated to the connection.
' dwReserved     [in] Reserved. Must be set to zero.
'
' Return:
' ¯¯¯¯¯¯¯
' Returns ERROR_SUCCESS if successful, or an error value otherwise. The error code can be one of the following:
'   ERROR_INVALID_PARAMETER One or more of the parameters are incorrect.
'   ERROR_NO_CONNECTION There is a problem with the dial-up connection.
'   ERROR_USER_DISCONNECTION The user clicked either the Work Offline or Cancel button on the Internet connection dialog box.
' ____________________________________________________________________________________________________________
' DWORD InternetDial (HWND hwndParent, LPTSTR lpszConnectoid, DWORD dwFlags, LPDWORD lpdwConnection, DWORD dwReserved);
'=============================================================================================================
Public Declare Function InternetDial Lib "wininet.dll" Alias "InternetDialA" (ByVal hWndParent As Long, ByVal lpszConnectoid As String, ByVal dwFlags As Long, ByRef lpdwConnection As Long, ByVal dwReserved As Long) As Long


'=============================================================================================================
' InternetErrorDlg
'
' Minimum Availability : Internet Explorer 3.0
'
' Parameters:
' ¯¯¯¯¯¯¯¯¯¯¯
' hWnd     [in]      Handle to the parent window for any needed dialog box. This parameter can be NULL if no dialog box is needed.
' hRequest [in, out] HINTERNET handle to the Internet connection used in the call to HttpSendRequest.
' dwError  [in]      Error value for which to display a dialog box. This can be one of the following values:
'   ERROR_INTERNET_HTTP_TO_HTTPS_ON_REDIR  Notifies the user of the zone crossing to and from a secure site.
'   ERROR_INTERNET_INCORRECT_PASSWORD      Displays a dialog box requesting the user's name and password. (On Microsoft® Windows® 95, the function attempts to use any cached authentication information for the server being accessed before displaying a dialog box.)
'   ERROR_INTERNET_INVALID_CA              Notifies the user that the Microsoft® Win32® Internet function does not recognize the certificate authority that generated the certificate for this Secure Sockets Layer (SSL) site.
'   ERROR_INTERNET_POST_IS_NON_SECURE      Displays a warning about posting data to the server through a nonsecure connection.
'   ERROR_INTERNET_SEC_CERT_CN_INVALID     Indicates that the SSL certificate Common Name (host name field) is incorrect. Displays an Invalid SSL Common Name dialog box and lets the user view the incorrect certificate. Also allows the user to select a certificate in response to a server request.
'   ERROR_INTERNET_SEC_CERT_DATE_INVALID   Tells the user that the SSL certificate has expired.
' dwFlags  [in]      Unsigned long integer value that contains the action flags. This can be a combination of these values: FLAGS_ERROR_UI_FILTER_FOR_ERRORS, FLAGS_ERROR_UI_FLAGS_CHANGE_OPTIONS, FLAGS_ERROR_UI_FLAGS_GENERATE_DATA, FLAGS_ERROR_UI_SERIALIZE_DIALOGS
' lppvData [in, out] Address of a pointer to a data structure. The structure can be different for each error that needs to be handled.
'
' Return:
' ¯¯¯¯¯¯¯
' Returns one of the following values, or an error value otherwise.
'   ERROR_SUCCESS The function completed successfully.
'   ERROR_CANCELLED The function was canceled by the user.
'   ERROR_INTERNET_FORCE_RETRY The Win32 function needs to redo its request.
'   ERROR_INVALID_HANDLE The handle to the parent window is invalid.
' ____________________________________________________________________________________________________________
' DWORD InternetErrorDlg (HWND hWnd, HINTERNET hRequest, DWORD dwError, DWORD dwFlags, LPVOID *lppvData);
'=============================================================================================================
Public Declare Function InternetErrorDlg Lib "wininet.dll" (ByVal hwnd As Long, ByRef hRequest As Long, ByVal dwError As Long, ByVal dwFlags As Long, ByRef lppvData As Any) As Long


'=============================================================================================================
' InternetFindNextFile
'
' Minimum Availability : Internet Explorer 3.0
'
' Parameters:
' ¯¯¯¯¯¯¯¯¯¯¯
' hFind       [in]  Valid HINTERNET handle returned from either FtpFindFirstFile or GopherFindFirstFile, or from InternetOpenUrl (directories only).
' lpvFindData [out] Address of the buffer that receives information about the found file or directory. The format of the information placed in the buffer depends on the protocol in use. The FTP protocol returns a WIN32_FIND_DATA  structure, and the Gopher protocol returns a GOPHER_FIND_DATA structure.
'
' Return:
' ¯¯¯¯¯¯¯
' Returns TRUE if the function succeeds, or FALSE otherwise. To get extended error information, call GetLastError . If the function finds no matching files, GetLastError returns ERROR_NO_MORE_FILES.
' ____________________________________________________________________________________________________________
' BOOL InternetFindNextFile (HINTERNET hFind, LPVOID lpvFindData);
'=============================================================================================================
Public Declare Function InternetFindNextFile_FTP Lib "wininet.dll" Alias "InternetFindNextFileA" (ByVal hFind As Long, ByRef lpvFindData As WIN32_FIND_DATA) As Long 'BOOL
Public Declare Function InternetFindNextFile_Gopher Lib "wininet.dll" Alias "InternetFindNextFileA" (ByVal hFind As Long, ByRef lpvFindData As GOPHER_FIND_DATA) As Long 'BOOL


'=============================================================================================================
' InternetGetConnectedState
'
' Minimum Availability : Internet Explorer 4.0
'
' Parameters:
' ¯¯¯¯¯¯¯¯¯¯¯
' lpdwFlags  [out] Address of an unsigned long integer variable where the connection description should be returned. This can be a combination of the following values: INTERNET_CONNECTION_CONFIGURED, INTERNET_CONNECTION_LAN, INTERNET_CONNECTION_MODEM, INTERNET_CONNECTION_MODEM_BUSY, INTERNET_CONNECTION_OFFLINE, INTERNET_CONNECTION_PROXY, INTERNET_RAS_INSTALLED
' dwReserved [in]  Reserved. Must be set to zero.
'
' Return:
' ¯¯¯¯¯¯¯
' Returns TRUE if there is an Internet connection, or FALSE otherwise.
' ____________________________________________________________________________________________________________
' BOOL InternetGetConnectedState (LPDWORD lpdwFlags, DWORD dwReserved);
'=============================================================================================================
Public Declare Function InternetGetConnectedState Lib "wininet.dll" (ByRef lpdwFlags As Long, ByVal dwReserved As Long) As Long 'BOOL


'=============================================================================================================
' InternetGetConnectedStateEx
'
' Minimum Availability : Internet Explorer 5
'
' Parameters:
' ¯¯¯¯¯¯¯¯¯¯¯
' lpdwFlags          [out] Address of an unsigned long integer variable where the connection description should be returned. This can be a combination of the following values: INTERNET_CONNECTION_CONFIGURED, INTERNET_CONNECTION_LAN, INTERNET_CONNECTION_MODEM, INTERNET_CONNECTION_MODEM_BUSY, INTERNET_CONNECTION_OFFLINE, INTERNET_CONNECTION_PROXY, INTERNET_RAS_INSTALLED
' lpszConnectionName [out] Address of a string value that receives the connection name.
' dwNameLen          [in]  Unsigned long integer value that contains the length of the lpszConnectionName string in TCHAR.
' dwReserved         [in]  Reserved. Must be set to zero.
'
' Return:
' ¯¯¯¯¯¯¯
' Returns TRUE if there is an Internet connection, or FALSE otherwise.
' ____________________________________________________________________________________________________________
' BOOL InternetGetConnectedStateEx (LPDWORD lpdwFlags, LPTSTR lpszConnectionName, DWORD dwNameLen, DWORD dwReserved);
'=============================================================================================================
Public Declare Function InternetGetConnectedStateEx Lib "wininet.dll" Alias "InternetGetConnectedStateExA" (ByRef lpdwFlags As Long, ByVal lpszConnectionName As String, ByVal dwNameLen As Long, ByVal dwReserved As Long) As Long 'BOOL


'=============================================================================================================
' InternetGetCookie
'
' Minimum Availability : Internet Explorer 3.0
'
' Parameters:
' ¯¯¯¯¯¯¯¯¯¯¯
' lpszUrlName    [in]      Address of a string that contains the URL to get cookies for.
' lpszCookieName [in]      Address of a string that contains the name of the cookie to get for the specified URL. This has not been implemented in this release.
' lpszCookieData [out]     Address of the buffer that receives the cookie data. This value can be NULL.
' lpdwSize       [in, out] Address of an unsigned long integer variable that specifies the size of the lpszCookieData buffer. If the function succeeds, the buffer receives the amount of data copied to the lpszCookieData buffer. If lpszCookieData is NULL, this parameter receives a value that specifies the size of the buffer necessary to copy all the cookie data.
'
' Return:
' ¯¯¯¯¯¯¯
' Returns TRUE if successful, or FALSE otherwise. To get a specific error value, call GetLastError . The following error values apply to InternetGetCookie:
'   ERROR_NO_MORE_ITEMS There is no cookie for the specified URL and all its parents.
'   ERROR_INSUFFICIENT_BUFFER The value passed in lpdwSize is insufficient to copy all the cookie data. The value returned in lpdwSize is the size of the buffer necessary to get all the data.
' ____________________________________________________________________________________________________________
' BOOL InternetGetCookie (LPCTSTR lpszUrlName, LPCTSTR lpszCookieName, LPTSTR lpszCookieData, LPDWORD lpdwSize);
'=============================================================================================================
Public Declare Function InternetGetCookie Lib "wininet.dll" Alias "InternetGetCookieA" (ByVal lpszUrlName As Long, ByVal lpszCookieName As String, ByVal lpszCookieData As String, ByRef lpdwSize As Long) As Long 'BOOL


'=============================================================================================================
' InternetGetLastResponseInfo
'
' Minimum Availability : Internet Explorer 3.0
'
' Parameters:
' ¯¯¯¯¯¯¯¯¯¯¯
' lpdwError        [out]     Address of an unsigned long integer variable that receives an error message pertaining to the operation that failed.
' lpszBuffer       [out]     Address of a buffer that receives the error text.
' lpdwBufferLength [in, out] Address of an unsigned long integer variable that contains the size of the lpszBuffer buffer in TCHARs. When the function returns, this parameter contains the size of the string written to the buffer, not including the terminating zero.
'
' Return:
' ¯¯¯¯¯¯¯
' Returns TRUE if error text was successfully written to the buffer, or FALSE otherwise. To get extended error information, call GetLastError . If the buffer is too small to hold all the error text, GetLastError returns ERROR_INSUFFICIENT_BUFFER, and the lpdwBufferLength parameter contains the minimum buffer size required to return all the error text.
' ____________________________________________________________________________________________________________
' BOOL InternetGetLastResponseInfo (LPDWORD lpdwError, LPTSTR lpszBuffer, LPDWORD lpdwBufferLength);
'=============================================================================================================
Public Declare Function InternetGetLastResponseInfo Lib "wininet.dll" Alias "InternetGetLastResponseInfoA" (ByRef lpdwError As Long, ByVal lpszBuffer As String, ByRef lpdwBufferLength As Long) As Long 'BOOL


'=============================================================================================================
' InternetGoOnline
'
' Minimum Availability : Internet Explorer 4.0
'
' Parameters:
' ¯¯¯¯¯¯¯¯¯¯¯
' lpszURL    [in] Address of a string variable containing the URL of the Web site to connect to.
' hwndParent [in] Handle to the parent window.
' dwReserved [in] Reserved. Must be set to zero.
'
' Return:
' ¯¯¯¯¯¯¯
' Returns TRUE if successful, or FALSE otherwise.
' ____________________________________________________________________________________________________________
' BOOL InternetGoOnline (LPTSTR lpszURL, HWND hwndParent, DWORD dwReserved);
'=============================================================================================================
Public Declare Function InternetGoOnline Lib "wininet.dll" Alias "InternetGoOnlineA" (ByVal lpszUrl As String, ByVal hWndParent As Long, ByVal dwReserved As Long) As Long 'BOOL


'=============================================================================================================
' InternetHangUp
'
' Minimum Availability : Internet Explorer 4.0
'
' Parameters:
' ¯¯¯¯¯¯¯¯¯¯¯
' dwConnection [in] Unsigned long integer value that contains the number assigned to the connection to be disconnected.
' dwReserved   [in] Reserved. Must be set to zero.
'
' Return:
' ¯¯¯¯¯¯¯
' Returns ERROR_SUCCESS if successful, or an error value otherwise.
' ____________________________________________________________________________________________________________
' DWORD InternetHangUp (DWORD dwConnection, DWORD dwReserved);
'=============================================================================================================
Public Declare Function InternetHangUp Lib "wininet.dll" (ByVal dwConnection As Long, ByVal dwReserved As Long) As Long


'=============================================================================================================
' InternetInitializeAutoProxyDll
'
' Minimum Availability  : Not currently supported.
'
' Parameters:
' ¯¯¯¯¯¯¯¯¯¯¯
' ?
' Return:
' ¯¯¯¯¯¯¯
' ?
' ____________________________________________________________________________________________________________
' ?
'=============================================================================================================


'=============================================================================================================
' InternetLockRequestFile
'
' Minimum Availability : Internet Explorer 3.0
'
' Parameters:
' ¯¯¯¯¯¯¯¯¯¯¯
' hInternet        [in]  HINTERNET handle returned by FtpOpenFile, GopherOpenFile, HttpOpenRequest, or InternetOpenUrl.
' lphLockReqHandle [out] Address of a handle to store the lock request handle.
'
' Return:
' ¯¯¯¯¯¯¯
' Returns TRUE if successful, or FALSE otherwise. To get a specific error message, call GetLastError .
' ____________________________________________________________________________________________________________
' BOOL InternetLockRequestFile (HINTERNET hInternet, HANDLE *lphLockReqHandle);
'=============================================================================================================
Public Declare Function InternetLockRequestFile Lib "wininet.dll" (ByVal hInternet As Long, ByRef lphLockReqHandle As Long) As Long


'=============================================================================================================
' InternetOpen
'
' Minimum Availability : Internet Explorer 3.0
'
' *NOTE - You can not use INTERNET_FLAG_ASYNC as a value for the "dwFlags" parameter using VB6 - it won't work.
'         You can, however, use it with VB5... VB5 allows it where VB6 does not.  Under VB6, specify ZERO (0)
'         for the dwFlags parameter.
'
' Parameters:
' ¯¯¯¯¯¯¯¯¯¯¯
' lpszAgent       [in] Address of a string variable that contains the name of the application or entity calling the Internet functions (for example, Microsoft® Internet Explorer). This name is used as the user agent in the HTTP protocol.
' dwAccessType    [in] Type of access required. This can be one of the following values: INTERNET_OPEN_TYPE_DIRECT, INTERNET_OPEN_TYPE_PRECONFIG, INTERNET_OPEN_TYPE_PRECONFIG_WITH_NO_AUTOPROXY, INTERNET_OPEN_TYPE_PROXY
' lpszProxyName   [in] Address of a string variable that contains the name of the proxy server(s) to use when proxy access is specified by setting dwAccessType to INTERNET_OPEN_TYPE_PROXY. Do not use an empty string, because InternetOpen will use it as the proxy name. The Win32 Internet functions recognize only CERN type proxies (HTTP only) and the TIS FTP gateway (FTP only). If Internet Explorer is installed, the Win32 Internet functions also support SOCKS proxies. FTP and Gopher requests can be made through a CERN type proxy either by changing them to an HTTP request or by using InternetOpenUrl. If dwAccessType is not set to INTERNET_OPEN_TYPE_PROXY, this parameter is ignored and should be set to NULL. For more information about listing proxy servers, see the Listing Proxy Servers section of Enabling Internet Functionality.
' lpszProxyBypass [in] Address of a string variable that contains an optional list of host names or IP addresses, or both, that should not be routed through the proxy when dwAccessType is set to INTERNET_OPEN_TYPE_PROXY. The list can contain wildcards. Do not use an empty string, because InternetOpen will use it as the proxy bypass list. If this parameter specifies the "<local>" macro as the only entry, the function bypasses any host name that does not contain a period. If dwAccessType is not set to INTERNET_OPEN_TYPE_PROXY, this parameter is ignored and should be set to NULL.
' dwFlags         [in] Unsigned long integer value that contains the flags that indicate various options affecting the behavior of the function. This can be a combination of these values: INTERNET_FLAG_ASYNC, INTERNET_FLAG_FROM_CACHE, INTERNET_FLAG_OFFLINE
'
' Return:
' ¯¯¯¯¯¯¯
' Returns a valid handle that the application passes to subsequent Win32 Internet functions. If InternetOpen fails, it returns NULL. To retrieve a specific error message, call GetLastError.
' ____________________________________________________________________________________________________________
' HINTERNET InternetOpen (LPCTSTR lpszAgent, DWORD dwAccessType, LPCTSTR lpszProxyName, LPCTSTR lpszProxyBypass, DWORD dwFlags);
'=============================================================================================================
Public Declare Function InternetOpen Lib "wininet.dll" Alias "InternetOpenA" (ByVal lpszAgent As String, ByVal dwAccessType As Long, ByVal lpszProxyName As String, ByVal lpszProxyBypass As String, ByVal dwFlags As Long) As Long


'=============================================================================================================
' InternetOpenUrl
'
' Minimum Availability : Internet Explorer 3.0
'
' Parameters:
' ¯¯¯¯¯¯¯¯¯¯¯
' hInternet       [in] HINTERNET handle to the current Internet session. The handle must have been returned by a previous call to InternetOpen.
' lpszURL         [in] Address of a string variable that contains the URL to begin reading. Only URLs beginning with ftp:, gopher:, http:, or https: are supported.
' lpszHeaders     [in] Address of a string variable that contains the headers to be sent to the HTTP server. (For more information, see the description of the lpszHeaders parameter in the HttpSendRequest function.)
' dwHeadersLength [in] Unsigned long integer value that contains the length, in TCHARs, of the additional headers. If this parameter is -1L and lpszHeaders is not NULL, lpszHeaders is assumed to be zero-terminated (ASCIIZ) and the length is calculated.
' dwFlags         [in] Unsigned long integer value that contains the API flags. This can be one of the following values:
'   INTERNET_FLAG_EXISTING_CONNECT
'   INTERNET_FLAG_HYPERLINK
'   INTERNET_FLAG_IGNORE_CERT_CN_INVALID
'   INTERNET_FLAG_IGNORE_CERT_DATE_INVALID
'   INTERNET_FLAG_IGNORE_REDIRECT_TO_HTTP
'   INTERNET_FLAG_IGNORE_REDIRECT_TO_HTTPS
'   INTERNET_FLAG_KEEP_CONNECTION
'   INTERNET_FLAG_NEED_FILE
'   INTERNET_FLAG_NO_AUTH
'   INTERNET_FLAG_NO_AUTO_REDIRECT
'   INTERNET_FLAG_NO_CACHE_WRITE
'   INTERNET_FLAG_NO_COOKIES
'   INTERNET_FLAG_NO_UI
'   INTERNET_FLAG_PASSIVE
'   INTERNET_FLAG_PRAGMA_NOCACHE
'   INTERNET_FLAG_RAW_DATA
'   INTERNET_FLAG_RELOAD
'   INTERNET_FLAG_RESYNCHRONIZE
'   INTERNET_FLAG_SECURE
' dwContext       [in] Address of an unsigned long integer value that contains the application-defined value that is passed, along with the returned handle, to any callback functions.
'
' Return:
' ¯¯¯¯¯¯¯
' Returns a valid handle to the FTP, Gopher, or HTTP URL if the connection is successfully established, or NULL if the connection fails. To retrieve a specific error message, call GetLastError . To determine why access to the service was denied, call InternetGetLastResponseInfo.
' ____________________________________________________________________________________________________________
' HINTERNET InternetOpenUrl (HINTERNET hInternet, LPCTSTR lpszUrl, LPCTSTR lpszHeaders, DWORD dwHeadersLength, DWORD dwFlags, DWORD_PTR dwContext);
'=============================================================================================================
Public Declare Function InternetOpenUrl Lib "wininet.dll" Alias "InternetOpenUrlA" (ByVal hInternet As Long, ByVal lpszUrl As String, ByVal lpszHeaders As String, ByVal dwHeadersLength As Long, ByVal dwFlags As Long, ByRef dwContext As Long) As Long


'=============================================================================================================
' InternetQueryDataAvailable
'
' Minimum Availability : Internet Explorer 3.0
'
' Parameters:
' ¯¯¯¯¯¯¯¯¯¯¯
' hFile                      [in]  Valid HINTERNET handle, as returned by InternetOpenUrl, FtpOpenFile, GopherOpenFile, or HttpOpenRequest.
' lpdwNumberOfBytesAvailable [out] Optional. Address of an unsigned long integer variable that receives the number of available bytes.
' dwFlags                    [in]  Reserved. Must be set to zero.
' dwContext                  [in]  Reserved. Must be set to zero.
'
' Return:
' ¯¯¯¯¯¯¯
' Returns TRUE if the function succeeds, or FALSE otherwise. To get extended error information, call GetLastError . If the function finds no matching files, GetLastError returns ERROR_NO_MORE_FILES.
' ____________________________________________________________________________________________________________
' BOOL InternetQueryDataAvailable (HINTERNET hFile, LPDWORD lpdwNumberOfBytesAvailable, DWORD dwFlags, DWORD dwContext);
'=============================================================================================================
Public Declare Function InternetQueryDataAvailable Lib "wininet" (ByVal hFile As Long, ByRef lpdwNumberOfBytesAvailable As Long, ByVal dwFlags As Long, ByVal dwContext As Long) As Long 'BOOL


'=============================================================================================================
' InternetQueryOption
'
' Minimum Availability : Internet Explorer 3.0
'
' Parameters:
' ¯¯¯¯¯¯¯¯¯¯¯
' hInternet        [in]      HINTERNET handle on which to query information.
' dwOption         [in]      Unsigned long integer value that contains the Internet option to query. This can be one of the Option Flags values (INTERNET_OPTION_*).
' lpBuffer         [out]     Address of a buffer that receives the option setting. (See the INTERNET_OPTION_* constants documentation above)  Strings returned by InternetQueryOption are globally allocated, so the calling application must globally free the string when it is finished using it.
' lpdwBufferLength [in, out] Address of an unsigned long integer variable that contains the length of lpBuffer, in TCHARs. When the function returns, the variable receives the length of the data placed into lpBuffer. If GetLastError  returns ERROR_INSUFFICIENT_BUFFER, this parameter receives the number of bytes required to hold the requested information.
'
' Return:
' ¯¯¯¯¯¯¯
' Returns TRUE if successful, or FALSE otherwise. To get a specific error message, call GetLastError.
' ____________________________________________________________________________________________________________
' BOOL InternetQueryOption (HINTERNET hInternet, DWORD dwOption, LPVOID lpBuffer, LPDWORD lpdwBufferLength);
'=============================================================================================================
Public Declare Function InternetQueryOption Lib "wininet.dll" Alias "InternetQueryOptionA" (ByVal hInternet As Long, ByVal dwOption As Long, ByRef lpBuffer As Any, ByRef lpdwBufferLength As Long) As Long 'BOOL


'=============================================================================================================
' InternetReadFile
'
' Minimum Availability : Internet Explorer 3.0
'
' Parameters:
' ¯¯¯¯¯¯¯¯¯¯¯
' hFile                 [in]  Valid HINTERNET handle returned from a previous call to InternetOpenUrl, FtpOpenFile, GopherOpenFile, or HttpOpenRequest.
' lpBuffer              [in]  Address of a buffer that receives the data read.
' dwNumberOfBytesToRead [in]  Unsigned long integer value that contains the number of bytes to read.
' lpdwNumberOfBytesRead [out] Address of an unsigned long integer variable that receives the number of bytes read. InternetReadFile sets this value to zero before doing any work or error checking.
'
' Return:
' ¯¯¯¯¯¯¯
' Returns TRUE if successful, or FALSE otherwise. To get extended error information, call GetLastError . An application can also use InternetGetLastResponseInfo when necessary.
' ____________________________________________________________________________________________________________
' BOOL InternetReadFile (HINTERNET hFile, LPVOID lpBuffer, DWORD dwNumberOfBytesToRead, LPDWORD lpdwNumberOfBytesRead);
'=============================================================================================================
Public Declare Function InternetReadFile Lib "wininet.dll" (ByVal hFile As Long, ByRef lpBuffer As Any, ByVal dwNumberOfBytesToRead As Long, ByRef lpdwNumberOfBytesRead As Long) As Long 'BOOL


'=============================================================================================================
' InternetReadFileEx
'
' Minimum Availability : Internet Explorer 4.0
'
' Parameters:
' ¯¯¯¯¯¯¯¯¯¯¯
' hFile        [in]  HINTERNET handle returned by the InternetOpenUrl or HttpOpenRequest function.
' lpBuffersOut [out] Address of an INTERNET_BUFFERS structure that contains the data downloaded.
' dwFlags      [in]  Unsigned long integer variable that contains the flags controlling the download. This can be the following value: IRF_NO_WAIT
' dwContext    [in]  Unsigned long integer variable that contains the context value used for asynchronous operations.
'
' Return:
' ¯¯¯¯¯¯¯
' Returns TRUE if successful, or FALSE otherwise. To get extended error information, call GetLastError . An application can also use InternetGetLastResponseInfo when necessary.
' ____________________________________________________________________________________________________________
' BOOL InternetReadFileEx (HINTERNET hFile, LPINTERNET_BUFFERS lpBuffersOut, DWORD dwFlags, DWORD dwContext);
'=============================================================================================================
Public Declare Function InternetReadFileEx Lib "wininet.dll" Alias "InternetReadFileExA" (ByVal hFile As Long, ByRef lpBuffersOut As INTERNET_BUFFERS, ByVal dwFlags As Long, ByVal dwContext As Long) As Long   'BOOL


'=============================================================================================================
' InternetSetCookie
'
' Minimum Availability : Internet Explorer 3.0
'
' Parameters:
' ¯¯¯¯¯¯¯¯¯¯¯
' lpszUrl        [in] Address of a null-terminated string that specifies the URL for which the cookie should be set.
' lpszCookieName [in] Address of a string that contains the name to associate with the cookie data. If this parameter is NULL, no name is associated with the cookie.
' lpszCookieData [in] Address of the actual data to associate with the URL.
'
' Return:
' ¯¯¯¯¯¯¯
' Returns TRUE if successful, or FALSE otherwise. To get a specific error message, call GetLastError .
' ____________________________________________________________________________________________________________
' BOOL InternetSetCookie (LPCTSTR lpszUrl, LPCTSTR lpszCookieName, LPCTSTR lpszCookieData);
'=============================================================================================================
Public Declare Function InternetSetCookie Lib "wininet.dll" Alias "InternetSetCookieA" (ByVal lpszUrl As String, ByVal lpszCookieName As String, ByVal lpszCookieData As Long) As Long 'BOOL


'=============================================================================================================
' InternetSetDialState
'
' Minimum Availability  : Not currently supported. This function is obsolete. Do not use.
'
' Parameters:
' ¯¯¯¯¯¯¯¯¯¯¯
' ?
' Return:
' ¯¯¯¯¯¯¯
' ?
' ____________________________________________________________________________________________________________
' ?
'=============================================================================================================


'=============================================================================================================
' InternetSetFilePointer
'
' Minimum Availability : Internet Explorer 3.0
'
' Parameters:
' ¯¯¯¯¯¯¯¯¯¯¯
' hFile           [in] Valid HINTERNET handle returned from a previous call to InternetOpenUrl (on an HTTP or HTTPS URL) or HttpOpenRequest (using the GET or HEAD method and passed to HttpSendRequest or HttpSendRequestEx). This handle must not have been created with the INTERNET_FLAG_DONT_CACHE or INTERNET_FLAG_NO_CACHE_WRITE value set.
' lDistanceToMove [in] A long integer value that contains the number of bytes to move the file pointer. A positive value moves the pointer forward in the file; a negative value moves it backward.
' pReserved       [in] Reserved. Must be set to NULL.
' dwMoveMethod    [in] Unsigned long integer value that indicates the starting point for the file pointer move. This can be one of the following values: FILE_BEGIN, FILE_CURRENT, FILE_END
' dwContext       [in] Reserved. Must be set to zero.
'
' Return:
' ¯¯¯¯¯¯¯
' Returns the current file position if the function succeeds, or -1 otherwise.
' ____________________________________________________________________________________________________________
' DWORD InternetSetFilePointer (HINTERNET hFile, LONG lDistanceToMove, PVOID pReserved, DWORD dwMoveMethod, DWORD dwContext);
'=============================================================================================================
Public Declare Function InternetSetFilePointer Lib "wininet.dll" (ByVal hFile As Long, ByVal lDistanceToMove As Long, ByVal pReserved As Long, ByVal dwMoveMethod As Long, ByVal dwContext As Long) As Long


'=============================================================================================================
' InternetSetOption
'
' Minimum Availability : Internet Explorer 3.0
'
' Parameters:
' ¯¯¯¯¯¯¯¯¯¯¯
' hInternet      [in] HINTERNET handle on which to set information.
' dwOption       [in] Unsigned long integer value that contains the Internet option to set. This can be one of the Option Flags values (INTERNET_OPTION_*).
' lpBuffer       [in] Address of a buffer that contains the option setting. (See the INTERNET_OPTION_* constants documentation above)
' dwBufferLength [in] Unsigned long integer value that contains the length of the lpBuffer buffer in TCHARs.
'
' Return:
' ¯¯¯¯¯¯¯
' Returns TRUE if successful, or FALSE otherwise. To get a specific error message, call GetLastError.
' ____________________________________________________________________________________________________________
' BOOL InternetSetOption (HINTERNET hInternet, DWORD dwOption, LPVOID lpBuffer, DWORD dwBufferLength);
'=============================================================================================================
Public Declare Function InternetSetOption Lib "wininet.dll" Alias "InternetSetOptionA" (ByVal hInternet As Long, ByVal dwOption As Long, ByRef lpBuffer As Any, ByVal dwBufferLength As Long) As Long


'=============================================================================================================
' InternetSetOptionEx
'
' Minimum Availability  : Not currently implemented.
'
' Parameters:
' ¯¯¯¯¯¯¯¯¯¯¯
' ?
' Return:
' ¯¯¯¯¯¯¯
' ?
' ____________________________________________________________________________________________________________
' ?
'=============================================================================================================


'=============================================================================================================
' InternetSetStatusCallback
'
' Minimum Availability : Internet Explorer 3.0
'
' Parameters:
' ¯¯¯¯¯¯¯¯¯¯¯
' hInternet            [in] HINTERNET handle for which the callback is to be set.
' lpfnInternetCallback [in] Address of the callback function to call when progress is made, or to return NULL to remove the existing callback function. For more information about the callback function, see INTERNET_STATUS_CALLBACK.
'
' Return:
' ¯¯¯¯¯¯¯
' Returns the previously defined status callback function if successful, NULL if there was no previously defined status callback function, or INTERNET_INVALID_STATUS_CALLBACK (-1) if the callback function is not valid.
' ____________________________________________________________________________________________________________
' INTERNET_STATUS_CALLBACK InternetSetStatusCallback (HINTERNET hInternet, INTERNET_STATUS_CALLBACK lpfnInternetCallback);
'=============================================================================================================
Public Declare Function InternetSetStatusCallback Lib "wininet.dll" Alias "InternetSetStatusCallbackA" (ByVal hInternet As Long, ByVal CallbackFunctionAddress As Long) As Long


'=============================================================================================================
' InternetTimeFromSystemTime
'
' Minimum Availability : Internet Explorer 3.0
'
' Parameters:
' ¯¯¯¯¯¯¯¯¯¯¯
' pst      [in]  Address of a SYSTEMTIME  structure that contains the date and time to format.
' dwRFC    [in]  Unsigned long integer value that contains the RFC format used. Currently, the only valid format is INTERNET_RFC1123_FORMAT (0).
' lpszTime [out] Address of a string buffer that receives the formatted date and time. The buffer should be of size INTERNET_RFC1123_BUFSIZE (30).
' cbTime   [in]  Unsigned long integer value that contains the size, in bytes, of the lpszTime buffer.
'
' Return:
' ¯¯¯¯¯¯¯
' Returns TRUE if the function succeeds, or FALSE otherwise. To get extended error information, call GetLastError.
' ____________________________________________________________________________________________________________
' BOOL InternetTimeFromSystemTime (const SYSTEMTIME *pst, DWORD dwRFC, LPTSTR lpszTime, DWORD cbTime);
'=============================================================================================================
Public Declare Function InternetTimeFromSystemTime Lib "wininet.dll" Alias "InternetTimeFromSystemTimeA" (ByRef pSystemTime As SYSTEMTIME, ByVal dwRFC As Long, ByVal lpszTime As String, ByVal cbTime As Long) As Long 'BOOL


'=============================================================================================================
' InternetTimeToSystemTime
'
' Minimum Availability : Internet Explorer 3.0
'
' Parameters:
' ¯¯¯¯¯¯¯¯¯¯¯
' lpszTime   [in]  Address of a null-terminated date/time string to convert.
' pst        [out] Address of SYSTEMTIME structure that receives the converted time.
' dwReserved [in]  Reserved. Must be set to zero.
'
' Return:
' ¯¯¯¯¯¯¯
' Returns TRUE if the string was converted, or FALSE otherwise. To get extended error information, call GetLastError .
' ____________________________________________________________________________________________________________
' BOOL InternetTimeToSystemTime (LPCTSTR lpszTime, SYSTEMTIME *pst, DWORD dwReserved);
'=============================================================================================================
Public Declare Function InternetTimeToSystemTime Lib "wininet.dll" Alias "InternetTimeToSystemTimeA" (ByVal lpszTime As String, ByRef pSystemTime As SYSTEMTIME, ByVal dwReserved As Long) As Long 'BOOL


'=============================================================================================================
' InternetUnlockRequestFile
'
' Minimum Availability : Internet Explorer 3.0
'
' Parameters:
' ¯¯¯¯¯¯¯¯¯¯¯
' hLockHandle [in] Lock request handle that was returned by InternetLockRequestFile.
'
' Return:
' ¯¯¯¯¯¯¯
' Returns TRUE if successful, or FALSE otherwise. To get a specific error message, call GetLastError.
' ____________________________________________________________________________________________________________
' BOOL InternetUnlockRequestFile (Handle hLockHandle);
'=============================================================================================================
Public Declare Function InternetUnlockRequestFile Lib "wininet.dll" (ByVal hLockHandle As Long) As Long 'BOOL


'=============================================================================================================
' InternetWriteFile
'
' Minimum Availability : Internet Explorer 3.0
'
' Parameters:
' ¯¯¯¯¯¯¯¯¯¯¯
' hFile                    [in]  Valid HINTERNET handle returned from a previous call to FtpOpenFile or an HINTERNET handle sent by HttpSendRequestEx.
' lpBuffer                 [in]  Address of a buffer that contains the data to be written to the file.
' dwNumberOfBytesToWrite   [in]  Unsigned long integer value that contains the number of bytes to write to the file.
' lpdwNumberOfBytesWritten [out] Address of an unsigned long integer variable that receives the number of bytes written to the buffer. InternetWriteFile sets this value to zero before doing any work or error checking.
'
' Return:
' ¯¯¯¯¯¯¯
' Returns TRUE if the function succeeds, or FALSE otherwise. To get extended error information, call GetLastError . An application can also use InternetGetLastResponseInfo when necessary.
' ____________________________________________________________________________________________________________
' BOOL InternetWriteFile (HINTERNET hFile, LPCVOID lpBuffer, DWORD dwNumberOfBytesToWrite, LPDWORD lpdwNumberOfBytesWritten);
'=============================================================================================================
Public Declare Function InternetWriteFile Lib "wininet.dll" (ByVal hFile As Long, ByRef lpBuffer As Any, ByVal dwNumberOfBytesToWrite As Long, ByRef lpdwNumberOfBytesWritten As Long) As Long  'BOOL


'=============================================================================================================
' ReadUrlCacheEntryStream
'
' Minimum Availability : Internet Explorer 3.0
'
' Parameters:
' ¯¯¯¯¯¯¯¯¯¯¯
' hUrlCacheStream [in]      HINTERNET handle that was returned by the RetrieveUrlCacheEntryStream function.
' dwLocation      [in]      Unsigned long integer value that contains the offset to read from.
' lpBuffer        [in, out] Address of a buffer that receives the data.
' lpdwLen         [in, out] Address of an unsigned long integer variable that specifies the length of the lpBuffer buffer, in TCHARs. When the function returns, the variable contains the number of TCHARs copied to the buffer, or the required size of the buffer, in bytes.
' dwReserved      [in]      Reserved. Must be set to zero.
'
' Return:
' ¯¯¯¯¯¯¯
' Returns TRUE if successful, or FALSE otherwise. To get extended error information, call GetLastError.
' ____________________________________________________________________________________________________________
' BOOL ReadUrlCacheEntryStream (HANDLE hUrlCacheStream, DWORD dwLocation, LPVOID lpBuffer, LPDWORD lpdwLen, DWORD dwReserved);
'=============================================================================================================
Public Declare Function ReadUrlCacheEntryStream Lib "wininet.dll" (ByVal hUrlCacheStream As Long, ByVal dwLocation As Long, ByRef lpBuffer As Any, ByRef lpdwLen As Long, ByVal dwReserved As Long) As Long 'BOOL


'=============================================================================================================
' RetrieveUrlCacheEntryFile
'
' Minimum Availability : Internet Explorer 3.0
'
' Parameters:
' ¯¯¯¯¯¯¯¯¯¯¯
' lpszUrlName                  [in]      Address of a string that contains the URL of the resource associated with the cache entry. This must be a unique name. The name string should not contain any escape characters.
' lpCacheEntryInfo             [out]     Address of a cache entry information buffer. If the buffer is not sufficient, this function returns ERROR_INSUFFICIENT_BUFFER and sets lpdwCacheEntryInfoBufferSize to the number of bytes required.
' lpdwCacheEntryInfoBufferSize [in, out] Address of an unsigned long integer variable that specifies the size of the lpCacheEntryInfo buffer, in TCHARs. When the function returns, the variable contains the size, in TCHARs, of the actual buffer used or the number of bytes required to retrieve the cache entry file. The caller should check the return value in this parameter. If the return size is less than or equal to the size passed in, all the relevant data has been returned.
' dwReserved                   [in]      Reserved. Must be set to zero.
'
' Return:
' ¯¯¯¯¯¯¯
' Returns TRUE if successful, or FALSE otherwise. To get extended error information, call GetLastError . Possible error values include:
'   ERROR_FILE_NOT_FOUND      The cache entry specified by the source name is not found in the cache storage.
'   ERROR_INSUFFICIENT_BUFFER The size of the lpCacheEntryInfo buffer as specified by lpdwCacheEntryInfoBufferSize is not sufficient to contain all the information. The value returned in lpdwCacheEntryInfoBufferSize indicates the buffer size necessary to get all the information.
' ____________________________________________________________________________________________________________
' BOOL RetrieveUrlCacheEntryFile (LPCTSTR lpszUrlName, LPINTERNET_CACHE_ENTRY_INFO lpCacheEntryInfo, LPDWORD lpdwCacheEntryInfoBufferSize, DWORD dwReserved);
'=============================================================================================================
Public Declare Function RetrieveUrlCacheEntryFile Lib "wininet.dll" Alias "RetrieveUrlCacheEntryFileA" (ByVal lpszUrlName As String, ByRef lpCacheEntryInfo As INTERNET_CACHE_ENTRY_INFO, ByRef lpdwCacheEntryInfoBufferSize As Long, ByVal dwReserved As Long) As Long 'BOOL


'=============================================================================================================
' RetrieveUrlCacheEntryStream
'
' Minimum Availability : Internet Explorer 3.0
'
' Parameters:
' ¯¯¯¯¯¯¯¯¯¯¯
' lpszUrlName                  [in]      Address of a string that contains the source name of the cache entry. This must be a unique name. The name string should not contain any escape characters.
' lpCacheEntryInfo             [out]     Address of an INTERNET_CACHE_ENTRY_INFO structure that receives information about the cache entry.
' lpdwCacheEntryInfoBufferSize [in, out] Address of an unsigned long integer variable that specifies the size of the lpCacheEntryInfo buffer, in TCHARs. When the function returns, the variable receives the number of TCHARs copied to the buffer, or the required size of the buffer, in bytes.
' fRandomRead                  [in]      BOOL (1/0) value that indicates whether the stream is open for random access. Set the flag to TRUE (1) to open the stream for random access.
' dwReserved                   [in]      Reserved. Must be set to zero.
'
' Return:
' ¯¯¯¯¯¯¯
' Returns a valid handle for use in the ReadUrlCacheEntryStream and UnlockUrlCacheEntryStream functions if successful, or NULL otherwise. To get extended error information, call GetLastError . Possible error values include:
'   ERROR_FILE_NOT_FOUND The cache entry specified by the source name is not found in the cache storage.
'   ERROR_INSUFFICIENT_BUFFER The size of lpCacheEntryInfo as specified by lpdwCacheEntryInfoBufferSize is not sufficient to contain all the information. The value returned in lpdwCacheEntryInfoBufferSize indicates the buffer size necessary to contain all the information.
' ____________________________________________________________________________________________________________
' HANDLE RetrieveUrlCacheEntryStream (LPCTSTR lpszUrlName, LPINTERNET_CACHE_ENTRY_INFO lpCacheEntryInfo, LPDWORD lpdwCacheEntryInfoBufferSize, BOOL fRandomRead, DWORD dwReserved);
'=============================================================================================================
Public Declare Function RetrieveUrlCacheEntryStream Lib "wininet.dll" Alias "RetrieveUrlCacheEntryStreamA" (ByVal lpszUrlName As String, ByRef lpCacheEntryInfo As INTERNET_CACHE_ENTRY_INFO, ByRef lpdwCacheEntryInfoBufferSize As Long, ByVal fRandomRead As Long, ByVal dwReserved As Long) As Long


'=============================================================================================================
' SetUrlCacheEntryGroup
'
' Minimum Availability : Internet Explorer 4.0
'
' Parameters:
' ¯¯¯¯¯¯¯¯¯¯¯
' lpszUrlName       [in] Address of a string value that contains the URL of the cached resource.
' dwFlags           [in] Unsigned long integer value that determines whether the entry is added to or removed from a cache group. This can be one of the following values: INTERNET_CACHE_GROUP_ADD, INTERNET_CACHE_GROUP_REMOVE
' GroupId           [in] GROUPID value that indicates the cache group that the entry will be added to or removed from.
' pbGroupAttributes [in] Reserved. Must be set to NULL.
' cbGroupAttributes [in] Reserved. Must be set to zero.
' lpReserved        [in] Reserved. Must be set to NULL.
'
' Return:
' ¯¯¯¯¯¯¯
' Returns TRUE if successful, or FALSE otherwise.
' ____________________________________________________________________________________________________________
' BOOL SetUrlCacheEntryGroup (LPCTSTR lpszUrlName, DWORD dwFlags, GROUPID GroupId, LPBYTE pbGroupAttributes, DWORD cbGroupAttributes, LPVOID lpReserved);
'=============================================================================================================
Public Declare Function SetUrlCacheEntryGroup Lib "wininet.dll" Alias "SetUrlCacheEntryGroupA" (ByVal lpszUrlName As String, ByVal dwFlags As Long, ByVal GroupId As Currency, ByVal pbGroupAttributes As Long, ByVal cbGroupAttributes As Long, ByVal lpReserved As Long) As Long


'=============================================================================================================
' SetUrlCacheEntryInfo
'
' Minimum Availability : Internet Explorer 3.0
'
' Parameters:
' ¯¯¯¯¯¯¯¯¯¯¯
' lpszUrlName      [in] Address of a string that contains the name of the cache entry. The name string should not contain any escape characters.
' lpCacheEntryInfo [in] Address of an INTERNET_CACHE_ENTRY_INFO structure containing the values to be assigned to the cache entry designated by lpszUrlName.
' dwFieldControl   [in] Unsigned long integer value that contains a bitmask that indicates the members that are to be set. This can be a combination of the following values: CACHE_ENTRY_ACCTIME_FC, CACHE_ENTRY_ATTRIBUTE_FC, CACHE_ENTRY_EXEMPT_DELTA_FC, CACHE_ENTRY_EXPTIME_FC, CACHE_ENTRY_HEADERINFO_FC, CACHE_ENTRY_HITRATE_FC, CACHE_ENTRY_MODTIME_FC, CACHE_ENTRY_SYNCTIME_FC
'
' Return:
' ¯¯¯¯¯¯¯
' Returns TRUE if successful, or FALSE otherwise. To get extended error information, call GetLastError . Possible error values include:
'   ERROR_FILE_NOT_FOUND    The specified cache entry is not found in the cache.
'   ERROR_INVALID_PARAMETER The value(s) to be set is invalid.
' ____________________________________________________________________________________________________________
' BOOL SetUrlCacheEntryInfo (LPCTSTR lpszUrlName, LPINTERNET_CACHE_ENTRY_INFO lpCacheEntryInfo, DWORD dwFieldControl);
'=============================================================================================================
Public Declare Function SetUrlCacheEntryInfo Lib "wininet.dll" Alias "SetUrlCacheEntryInfoA" (ByVal lpszUrlName As String, ByRef lpCacheEntryInfo As INTERNET_CACHE_ENTRY_INFO, ByVal dwFieldControl As Long) As Long


'=============================================================================================================
' SetUrlCacheGroupAttribute
'
' Minimum Availability : Internet Explorer 5
'
' Parameters:
' ¯¯¯¯¯¯¯¯¯¯¯
' gID          [in]      GROUPID of the cache group.
' dwFlags      [in]      Reserved. Must be set to zero.
' dwAttributes [in]      Unsigned long integer value that indicates what attributes to set. This can be one of the following values: CACHEGROUP_ATTRIBUTE_FLAG, CACHEGROUP_ATTRIBUTE_GROUPNAME, CACHEGROUP_ATTRIBUTE_QUOTA, CACHEGROUP_ATTRIBUTE_STORAGE, CACHEGROUP_ATTRIBUTE_TYPE, CACHEGROUP_READWRITE_MASK
' lpGroupInfo  [in]      Address of an INTERNET_CACHE_GROUP_INFO structure that contains the attribute information to store.
' lpReserved   [in, out] Reserved. Must be set to NULL.
'
' Return:
' ¯¯¯¯¯¯¯
' Returns TRUE if successful, or FALSE otherwise. To get specific error information, call GetLastError.
' ____________________________________________________________________________________________________________
' BOOL SetUrlCacheGroupAttribute (GROUPID gID, DWORD dwFlags, DWORD dwAttributes, LPINTERNET_CACHE_GROUP_INFO lpGroupInfo, LPVOID lpReserved);
'=============================================================================================================
Public Declare Function SetUrlCacheGroupAttribute Lib "wininet.dll" (ByVal gID As Currency, ByVal dwFlags As Long, ByVal dwAttributes As Long, ByRef lpGroupInfo As INTERNET_CACHE_ENTRY_INFO, ByRef lpReserved As Long) As Long


'=============================================================================================================
' UnlockUrlCacheEntryFile
'
' Minimum Availability : Internet Explorer 3.0
'
' Parameters:
' ¯¯¯¯¯¯¯¯¯¯¯
' lpszUrlName [in] Address of a string that contains the source name of the cache entry that is being unlocked. The name string should not contain any escape characters.
' dwReserved  [in] Reserved. Must be set to zero.
'
' Return:
' ¯¯¯¯¯¯¯
' Returns TRUE if successful, or FALSE otherwise. To get extended error information, call GetLastError . ERROR_FILE_NOT_FOUND indicates that the cache entry specified by the source name is not found in the cache storage.
' ____________________________________________________________________________________________________________
' BOOL UnlockUrlCacheEntryFile (LPCTSTR lpszUrlName, DWORD dwReserved);
'=============================================================================================================
Public Declare Function UnlockUrlCacheEntryFile Lib "wininet.dll" Alias "UnlockUrlCacheEntryFileA" (ByVal lpszUrlName As String, ByVal dwReserved As Long) As Long


'=============================================================================================================
' UnlockUrlCacheEntryStream
'
' Minimum Availability : Internet Explorer 3.0
'
' Parameters:
' ¯¯¯¯¯¯¯¯¯¯¯
' hUrlCacheStream [in] Handle that was returned by the RetrieveUrlCacheEntryStream function.
' dwReserved      [in] Reserved. Must be set to zero.
'
' Return:
' ¯¯¯¯¯¯¯
' Returns TRUE if successful, or FALSE otherwise. To get extended error information, call GetLastError .
' ____________________________________________________________________________________________________________
' BOOL UnlockUrlCacheEntryStream (HANDLE hUrlCacheStream, DWORD dwReserved);
'=============================================================================================================
Public Declare Function UnlockUrlCacheEntryStream Lib "wininet.dll" (ByVal hUrlCacheStream As Long, ByVal dwReserved As Long) As Long





'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX



'=============================================================================================================
' Prototype for an application-defined status callback function.
'
' typedef void (CALLBACK *INTERNET_STATUS_CALLBACK)(
'   HINTERNET  hInternet,                 // [in] Handle for which the callback function is being called.
'   DWORD_PTR  dwContext,                 // [in] Address of an unsigned long integer value that contains the application-defined context value associated with hInternet.
'   DWORD      dwInternetStatus,          // [in] Unsigned long integer value that contains the status code that indicates why the callback function is being called. This can be one of the following values: INTERNET_STATUS_* (See Constants Above)
'   LPVOID     lpvStatusInformation,      // [in] Address of a buffer that contains information pertinent to this call to the callback function. (This information changes depending on what the dwInternetStatus is... see constants documentation above)
'   DWORD      dwStatusInformationLength  // [in] Unsigned long integer value that contains the size, in TCHARs, of the lpvStatusInformation buffer.
' );
'=============================================================================================================
Public Sub CallbackProc(ByVal hInternet As Long, ByVal dwContext As Long, ByVal dwInternetStatus As Long, ByRef lpvStatusInformation As Long, ByVal dwStatusInformationLength As Long)
  
  Dim SA_Temp       As sockaddr
  Dim IAR_Temp      As INTERNET_ASYNC_RESULT
  Dim RC_Temp       As REQUEST_CONTEXT
  Dim Str_Temp      As String
  Dim Lng_Temp      As Long
  Dim strStatusNum  As String
  Dim strStatusDesc As String
  Dim strStatusInfo As String
  Dim strIpAddress  As String
  Dim lngResultCode As Long
  Dim lngErrorNum   As Long
  Dim lngBytesSent  As Long
  Dim lngBytesRecv  As Long
  
  ' Get the status number
  strStatusNum = dwInternetStatus
  
  ' Get the status description
  strStatusDesc = CallbackStatus(dwInternetStatus)
  
  ' Get the Status Information
  Select Case dwInternetStatus
    Case INTERNET_STATUS_CONNECTED_TO_SERVER, INTERNET_STATUS_CONNECTING_TO_SERVER
      Str_Temp = String(MAX_PATH, Chr(0))
      StringFromPointer Str_Temp, VarPtr(lpvStatusInformation)
      Str_Temp = Left(Str_Temp, InStr(Str_Temp, Chr(0)) - 1)
      strStatusInfo = "IP Address = " & Str_Temp
      strIpAddress = Str_Temp
      
    Case INTERNET_STATUS_NAME_RESOLVED
      Str_Temp = String(MAX_PATH, Chr(0))
      StringFromPointer Str_Temp, VarPtr(lpvStatusInformation)
      Str_Temp = Left(Str_Temp, InStr(Str_Temp, Chr(0)) - 1)
      strStatusInfo = "IP Address = " & Str_Temp
      strIpAddress = Str_Temp
      
    Case INTERNET_STATUS_REDIRECT, INTERNET_STATUS_RESOLVING_NAME
      Str_Temp = String(MAX_PATH, Chr(0))
      StringFromPointer Str_Temp, VarPtr(lpvStatusInformation)
      Str_Temp = Left(Str_Temp, InStr(Str_Temp, Chr(0)) - 1)
      strStatusInfo = "IP Address = " & Str_Temp
      strIpAddress = Str_Temp
      
    Case INTERNET_STATUS_REQUEST_SENT
      strStatusInfo = "Bytes SENT = " & CStr(lpvStatusInformation)
      lngBytesSent = lpvStatusInformation
      
    Case INTERNET_STATUS_RESPONSE_RECEIVED
      strStatusInfo = "Byte RECIEVED = " & CStr(lpvStatusInformation)
      lngBytesRecv = lpvStatusInformation
      
    Case INTERNET_STATUS_HANDLE_CREATED, INTERNET_STATUS_REQUEST_COMPLETE
      CopyMemory IAR_Temp, VarPtr(lpvStatusInformation), Len(IAR_Temp)
      strStatusInfo = "Result = " & IAR_Temp.dwResult & ", Error = " & IAR_Temp.dwError
      lngResultCode = IAR_Temp.dwResult
      lngErrorNum = IAR_Temp.dwError
      
    Case Else
      strStatusInfo = ""
  End Select
  
End Sub

'=============================================================================================================
' CallbackStatus
'
' Purpose :
' Used to translate the "dwInternetStatus" parameter of the CallbackProc into a readable status message.
'
' Param                 Use
' ------------------------------------
' dwInternetStatus      This is the value passed to the CallbackProc by the "dwInternetStatus" parameter
'
' Return
' ------
' Returns the equivelant description for the status number
'
'=============================================================================================================
Public Function CallbackStatus(ByVal dwInternetStatus As Long) As String
  
  Select Case dwInternetStatus
    Case INTERNET_STATUS_CLOSING_CONNECTION    '50
      CallbackStatus = "Server connection closed"      '"Closing the connection to the server. The lpvStatusInformation parameter is NULL."
    Case INTERNET_STATUS_CONNECTED_TO_SERVER   '21
      CallbackStatus = "Connected successfully"     '"Successfully connected to the socket address (SOCKADDR) pointed to by lpvStatusInformation."
    Case INTERNET_STATUS_CONNECTING_TO_SERVER  '20
      CallbackStatus = "Connecting to socket"       '"Connecting to the socket address (SOCKADDR) pointed to by lpvStatusInformation."
    Case INTERNET_STATUS_CONNECTION_CLOSED     '51
      CallbackStatus = "Connection closed"             '"Successfully closed the connection to the server. The lpvStatusInformation parameter is NULL."
    Case INTERNET_STATUS_CTL_RESPONSE_RECEIVED '42
      CallbackStatus = ""                              ' < Not implemented >
    Case INTERNET_STATUS_DETECTING_PROXY       '80
      CallbackStatus = "Proxy detected"                '"Notifies the client application that a proxy has been detected."
    Case INTERNET_STATUS_HANDLE_CLOSING        '70
      CallbackStatus = "Connection handle terminated"  '"This handle value has been terminated."
    Case INTERNET_STATUS_HANDLE_CREATED        '60
      CallbackStatus = "New connection handle created" '"Used by InternetConnect to indicate it has created the new handle. This lets the application call InternetCloseHandle from another thread, if the connect is taking too long. The lpvStatusInformation parameter contains the address of an INTERNET_ASYNC_RESULT structure."
    Case INTERNET_STATUS_INTERMEDIATE_RESPONSE '120
      CallbackStatus = "Recieved intermediate status"  '"Received an intermediate (100 level) status code message from the server."
    Case INTERNET_STATUS_NAME_RESOLVED         '11
      CallbackStatus = "IP address found"              '"Successfully found the IP address of the name contained in lpvStatusInformation."
    Case INTERNET_STATUS_PREFETCH              '43
      CallbackStatus = ""                              ' < Not implemented >
    Case INTERNET_STATUS_RECEIVING_RESPONSE    '40
      CallbackStatus = "Waiting for server response"   '"Waiting for the server to respond to a request. The lpvStatusInformation parameter is NULL."
    Case INTERNET_STATUS_REDIRECT              '110
      CallbackStatus = "HTTP redirected"               '"An HTTP request is about to automatically redirect the request. The lpvStatusInformation parameter points to the new URL. At this point, the application can read any data returned by the server with the redirect response and can query the response headers. It can also cancel the operation by closing the handle. This callback is not made if the original request specified INTERNET_FLAG_NO_AUTO_REDIRECT."
    Case INTERNET_STATUS_REQUEST_COMPLETE      '100
      CallbackStatus = "Async operation completed"     '"An asynchronous operation has been completed. The lpvStatusInformation parameter contains the address of an INTERNET_ASYNC_RESULT structure."
    Case INTERNET_STATUS_REQUEST_SENT          '31
      CallbackStatus = "Server request sent"           '"Successfully sent the information request to the server. The lpvStatusInformation parameter points to a DWORD containing the number of bytes sent."
    Case INTERNET_STATUS_RESOLVING_NAME        '10
      CallbackStatus = "Resolving IP address"          '"Looking up the IP address of the name contained in lpvStatusInformation."
    Case INTERNET_STATUS_RESPONSE_RECEIVED     '41
      CallbackStatus = "Server response recieved"      '"Successfully received a response from the server. The lpvStatusInformation parameter points to a DWORD containing the number of bytes received."
    Case INTERNET_STATUS_SENDING_REQUEST       '30
      CallbackStatus = "Sending server request"        '"Sending the information request to the server. The lpvStatusInformation parameter is NULL."
    Case INTERNET_STATUS_STATE_CHANGE          '200
      CallbackStatus = "Moving between secure and nonsecure site" '"Moved between a secure (HTTPS) and a nonsecure (HTTP) site. This can be one of the following values:"
    Case INTERNET_STATE_CONNECTED              '&H1
      CallbackStatus = "Connected"                     '"Connected state (mutually exclusive with disconnected state)."
    Case INTERNET_STATE_DISCONNECTED           '&H2
      CallbackStatus = "No network connection could be established" '"Disconnected state. No network connection could be established."
    Case INTERNET_STATE_DISCONNECTED_BY_USER   '&H10
      CallbackStatus = "User disconnected"             '"Disconnected by user request."
    Case INTERNET_STATE_IDLE                   '&H100
      CallbackStatus = "Connection idle"               '"No network requests are being made by the Microsoft® Win32® Internet functions."
    Case INTERNET_STATE_BUSY                   '&H200
      CallbackStatus = "Connection busy"               '"Network requests are being made by the Win32 Internet functions."
    Case INTERNET_STATUS_USER_INPUT_REQUIRED   '140
      CallbackStatus = "User input required"           '"The request requires user input to be completed."
    Case Else
      CallbackStatus = "Unknown Status" ' <OTHER>
  End Select

End Function

'=============================================================================================================
' FiletimeToDate
'
' Purpose :
' This function takes the time/date of files returned by the WININET.DLL functions (UTC time) and converts it
' to the VB data type "Date" to be more easily used within VB.
'
' Param                 Use
' ------------------------------------
' Win32Time             The time/date of a file returned by the Win32 API (FILETIME structure which is
'                       represented by the VB data type "Currency")
' DisplayErrorMessages  Optional. If set to TRUE and an error occurs, an error message will pop up
'
' Return
' ------
' Returns TRUE if the function succeeds
' Returns FALSE if the function fails
'
'=============================================================================================================
Public Function FiletimeToDate(ByRef Win32Time As Currency, _
                               Optional ByVal DisplayErrorMessages As Boolean = False, _
                               Optional ByRef Return_ErrNum As Long, _
                               Optional ByRef Return_ErrSrc As String, _
                               Optional ByRef Return_ErrDesc As String) As Date
On Error GoTo ErrorTrap
   
  Dim LocalFileTime As Currency
  
  ' Make sure there is a valid FILETIME structure passed
  If Win32Time = 0 Then Exit Function
  
  If FileTimeToLocalFileTime(Win32Time, LocalFileTime) Then
    ' Local time is nanoseconds since 01-01-1601 in Currency that comes out as milliseconds.
    ' Divide by milliseconds per day to get days since 1601, then subtract days from 1601 to 1899 to get VB Date equivalent.
    FiletimeToDate = CDate((LocalFileTime / TIME_MILLISEC_PER_DAY) - TIME_DAY_ZERO)
  Else
    GetLastErrMsg Err.LastDllError, "FileTimeToLocalFileTime", Return_ErrNum, Return_ErrDesc, DisplayErrorMessages
    Return_ErrSrc = "FiletimeToDate >> FileTimeToLocalFileTime(..)"
    Err.Raise Return_ErrNum, Return_ErrSrc, Return_ErrDesc
  End If
  
  Exit Function
  
ErrorTrap:
  
  Return_ErrNum = Err.number
  Return_ErrSrc = Err.Source
  Return_ErrDesc = Err.Description
  Err.Clear
  
End Function

'=============================================================================================================
' CallbackSetup
'
' Purpose :
' This function tells the WININET.DLL to start sending callback messages to the CallbackProc function to
' let this module know what's going on within the functions of WININET.DLL.
'
' Param                 Use
' ------------------------------------
' hInternetSession      The handle of the internet connection to receive callback messages from
'                       (returned by the InternetOpen API)
' DisplayErrorMessages  Optional. If set to TRUE and an error occurs, an error message will pop up
' Return_ErrNum         Optional. If an error occurs, this variable returns the number of the error.
' Return_ErrSrc         Optional. If an error occurs, this variable returns the source of the error.
' Return_ErrDesc        Optional. If an error occurs, this variable returns the description of the error.
'
' Return
' ------
' Returns TRUE if the function succeeds
' Returns FALSE if the function fails
'
'=============================================================================================================
Public Function CallbackSetup(ByVal hInternetSession As Long, _
                              Optional ByVal DisplayErrorMessages As Boolean = False, _
                              Optional ByRef Return_ErrNum As Long, _
                              Optional ByRef Return_ErrSrc As String, _
                              Optional ByRef Return_ErrDesc As String) As Boolean
On Error GoTo ErrorTrap
  
  Dim FuncAddr As Long
  
  Return_ErrNum = 0
  Return_ErrSrc = ""
  Return_ErrDesc = ""
  
  ' Get the address of the Callback Procedure... this is required because we test the
  ' negative value of the Callback Procedure to see if the "InternetSetStatusCallback"
  ' API succeeded.  (See documentation for InternetSetStatusCallback)
  FuncAddr = GetFunctionAddress(AddressOf CallbackProc)
  
  ' Get the address of the previous callback procedure
  PrevCallbackAddr = InternetSetStatusCallback(hInternetSession, FuncAddr)
  If PrevCallbackAddr = (FuncAddr * -1) Then
    GetLastErrMsg Err.LastDllError, "InternetSetStatusCallback", Return_ErrNum, Return_ErrDesc, DisplayErrorMessages
    Return_ErrSrc = "CallbackSetup >> InternetSetStatusCallback(..)"
    Err.Raise Return_ErrNum, Return_ErrSrc, Return_ErrDesc
  Else
    CallbackSetup = True
  End If
  
  Exit Function
  
ErrorTrap:
  
  Return_ErrNum = Err.number
  Return_ErrSrc = Err.Source
  Return_ErrDesc = Err.Description
  Err.Clear
  
End Function

'=============================================================================================================
' CallbackReset
'
' Purpose :
' This function is the opposite of the "CallbackSetup" function where CallbackSetup sets up the CallbackProc,
' this function stops the callbacks to the CallbackProc.
'
' Param                 Use
' ------------------------------------
' hInternetSession      The handle to the internet connection to be reset (returned by the InternetOpen API)
' DisplayErrorMessages  Optional. If set to TRUE and an error occurs, an error message will pop up
' Return_ErrNum         Optional. If an error occurs, this variable returns the number of the error.
' Return_ErrSrc         Optional. If an error occurs, this variable returns the source of the error.
' Return_ErrDesc        Optional. If an error occurs, this variable returns the description of the error.
'
' Return
' ------
' Returns TRUE if the function succeeds
' Returns FALSE if the function fails
'
'=============================================================================================================
Public Function CallbackReset(ByVal hInternetSession As Long, _
                              Optional ByVal DisplayErrorMessages As Boolean = False, _
                              Optional ByRef Return_ErrNum As Long, _
                              Optional ByRef Return_ErrSrc As String, _
                              Optional ByRef Return_ErrDesc As String) As Boolean
On Error GoTo ErrorTrap
   
  Dim PrevAddr As Long
  
  Return_ErrNum = 0
  Return_ErrSrc = ""
  Return_ErrDesc = ""
  
  ' Set the callback procedure to what it was before
  PrevAddr = InternetSetStatusCallback(hInternetSession, PrevCallbackAddr)
  If PrevAddr <> (PrevCallbackAddr * -1) Then
    CallbackReset = True
    PrevCallbackAddr = 0
  End If
  
  ' Check for an error
  If GetLastErrMsg(Err.LastDllError, "InternetSetStatusCallback", Return_ErrNum, Return_ErrDesc, DisplayErrorMessages) = True Then
    CallbackReset = False
    Return_ErrSrc = "CallbackReset >> InternetSetStatusCallback(..)"
    Err.Raise Return_ErrNum, Return_ErrSrc, Return_ErrDesc
  Else
    CallbackReset = True
  End If
  
  Exit Function
  
ErrorTrap:
  
  Return_ErrNum = Err.number
  Return_ErrSrc = Err.Source
  Return_ErrDesc = Err.Description
  Err.Clear
  
End Function

'=============================================================================================================
' GetFunctionAddress
'
' Purpose :
' Simple function that takes the value passed by the "AddressOf" operator and returns it as a long so it can
' be used and compared.
'
' Param                 Use
' ------------------------------------
' AddressOfReturn       The address of the specified function returned by the "AddressOf" operator
'
' Return
' ------
' Returns the address as a long
'
'=============================================================================================================
Private Function GetFunctionAddress(ByVal AddressOfReturn As Long) As Long
  
  GetFunctionAddress = AddressOfReturn
  
End Function

'=============================================================================================================
' GetLastErrMsg
'
' Purpose :
' Function that gets the last error caused by Windows API's.  This only works with functions that use the
' GetLastError function to return an error code.  Not all Windows API's do.
'
' If no error has occured, no message is displayed.
'
' Param                 Use
' ------------------------------------
' ErrorNumber           Optional. Error number to display.  If this is set to zero, then the GetLastError
'                       API is called to see if any errors have occured.  If no error have occured, the
'                       function exits.
' LastAPICalled         Optional. If the "DisplayErrorMessage" parameter is set to TRUE, this is used to
'                       display an error dialog and tell the user which API caused the problem.
' Return_ErrNum         Optional. This returns the number of the error that just occured.  If the
'                       "ErrorNumber" parameter wasn't specified but an error was found by this function,
'                       the error number is returned here.
' Return_ErrDesc        Optional. This returns the description of the last error that occured.
' DisplayErrorMessage   Optional. If this parameter is set to TRUE, an error dialog is displayed with the
'                       error number and description for the user to see.
'
' Return
' ------
' If no error occured, no message is displayed & function returns FALSE.
' If an error occured, an error message is displayed & the function returns TRUE.
'
'=============================================================================================================
Public Function GetLastErrMsg(Optional ByVal ErrorNumber As Long, _
                              Optional ByVal LastAPICalled As String = "last", _
                              Optional ByRef Return_ErrNum As Long, _
                              Optional ByRef Return_ErrDesc As String, _
                              Optional ByVal DisplayErrorMessage As Boolean = True) As Boolean
On Error Resume Next
  
  Dim ExtendedErr As String
  Dim BufferLen   As Long
  Dim errnum      As Long
  
  ' Clear the return values first
  Return_ErrNum = 0
  Return_ErrDesc = ""
  
  ' If no error message is specified then check for one
  If ErrorNumber = 0 Then
    ErrorNumber = GetLastError
    If ErrorNumber = 0 Then
      GetLastErrMsg = False
      Exit Function
    End If
  End If
  
  ' Allocate a buffer for the error description
  Return_ErrNum = ErrorNumber
  Return_ErrDesc = String(MAX_PATH, Chr(0))
  BufferLen = MAX_PATH
  errnum = ErrorNumber
  
  ' Get the error description
  If ErrorNumber >= INTERNET_ERROR_FIRST And ErrorNumber <= INTERNET_ERROR_LAST Then
    InternetGetLastResponseInfo errnum, Return_ErrDesc, BufferLen
    Return_ErrDesc = Left(Return_ErrDesc, InStr(Return_ErrDesc, Chr(0)) - 1)
    If Right(Return_ErrDesc, Len(vbCrLf)) = vbCrLf Then
      Return_ErrDesc = Left(Return_ErrDesc, Len(Return_ErrDesc) - Len(vbCrLf))
    End If
    If Trim(Return_ErrDesc) = "" Then Return_ErrDesc = GetFtpErrorDescription(Return_ErrNum)
  Else
    FormatMessage FORMAT_MESSAGE_FROM_SYSTEM, ByVal 0&, errnum, 0, Return_ErrDesc, BufferLen, 0
    Return_ErrDesc = Left(Return_ErrDesc, InStr(Return_ErrDesc, Chr(0)) - 1)
    If Right(Return_ErrDesc, Len(vbCrLf)) = vbCrLf Then
      Return_ErrDesc = Left(Return_ErrDesc, Len(Return_ErrDesc) - Len(vbCrLf))
    End If
  End If
  
  ' Display the error message
  If DisplayErrorMessage = True Then MsgBox "An error occured while calling the " & LastAPICalled & " Windows API function." & Chr(13) & "Below is the error information:" & Chr(13) & Chr(13) & "Error Number = " & CStr(Return_ErrNum) & Chr(13) & "Error Description = " & Return_ErrDesc, vbOKOnly + vbExclamation, "  Windows API Error"
  GetLastErrMsg = True
  
  ' Set the last error to 0 (no error) so next time through it doesn't report the same error twice
  SetLastError 0
  
End Function

Public Function GetFtpErrorDescription(ByVal lngFtpErrorNumber As Long) As String
  
  Select Case lngFtpErrorNumber
    
    ' Internet API Error Returns
    Case ERROR_INTERNET_OUT_OF_HANDLES:           GetFtpErrorDescription = "No more handles could be generated at this time."
    Case ERROR_INTERNET_TIMEOUT:                  GetFtpErrorDescription = "The request has timed out."
    Case ERROR_INTERNET_EXTENDED_ERROR:           GetFtpErrorDescription = "An extended error was returned from the server. This is typically a string or buffer containing a verbose error message. Call InternetGetLastResponseInfo to retrieve the error text."
    Case ERROR_INTERNET_INTERNAL_ERROR:           GetFtpErrorDescription = "An internal error has occurred."
    Case ERROR_INTERNET_INVALID_URL:              GetFtpErrorDescription = "The URL is invalid."
    Case ERROR_INTERNET_UNRECOGNIZED_SCHEME:      GetFtpErrorDescription = "The URL scheme could not be recognized, or is not supported."
    Case ERROR_INTERNET_NAME_NOT_RESOLVED:        GetFtpErrorDescription = "The server name could not be resolved."
    Case ERROR_INTERNET_PROTOCOL_NOT_FOUND:       GetFtpErrorDescription = "The requested protocol could not be located."
    Case ERROR_INTERNET_INVALID_OPTION:           GetFtpErrorDescription = "A request to InternetQueryOption or InternetSetOption specified an invalid option value."
    Case ERROR_INTERNET_BAD_OPTION_LENGTH:        GetFtpErrorDescription = "The length of an option supplied to InternetQueryOption or InternetSetOption is incorrect for the type of option specified."
    Case ERROR_INTERNET_OPTION_NOT_SETTABLE:      GetFtpErrorDescription = "The requested option cannot be set, only queried."
    Case ERROR_INTERNET_SHUTDOWN:                 GetFtpErrorDescription = "The Win32 Internet function support is being shut down or unloaded."
    Case ERROR_INTERNET_INCORRECT_USER_NAME:      GetFtpErrorDescription = "The request to connect and log on to an FTP server could not be completed because the supplied user name is incorrect."
    Case ERROR_INTERNET_INCORRECT_PASSWORD:       GetFtpErrorDescription = "The request to connect and log on to an FTP server could not be completed because the supplied password is incorrect."
    Case ERROR_INTERNET_LOGIN_FAILURE:            GetFtpErrorDescription = "The request to connect and log on to an FTP server failed."
    Case ERROR_INTERNET_INVALID_OPERATION:        GetFtpErrorDescription = "The requested operation is invalid."
    Case ERROR_INTERNET_OPERATION_CANCELLED:      GetFtpErrorDescription = "The operation was canceled, usually because the handle on which the request was operating was closed before the operation completed."
    Case ERROR_INTERNET_INCORRECT_HANDLE_TYPE:    GetFtpErrorDescription = "The type of handle supplied is incorrect for this operation."
    Case ERROR_INTERNET_INCORRECT_HANDLE_STATE:   GetFtpErrorDescription = "The requested operation cannot be carried out because the handle supplied is not in the correct state."
    Case ERROR_INTERNET_NOT_PROXY_REQUEST:        GetFtpErrorDescription = "The request cannot be made via a proxy."
    Case ERROR_INTERNET_REGISTRY_VALUE_NOT_FOUND: GetFtpErrorDescription = "A required registry value could not be located."
    Case ERROR_INTERNET_BAD_REGISTRY_PARAMETER:   GetFtpErrorDescription = "A required registry value was located but is an incorrect type or has an invalid value."
    Case ERROR_INTERNET_NO_DIRECT_ACCESS:         GetFtpErrorDescription = "Direct network access cannot be made at this time."
    Case ERROR_INTERNET_NO_CONTEXT:               GetFtpErrorDescription = "An asynchronous request could not be made because a zero context value was supplied."
    Case ERROR_INTERNET_NO_CALLBACK:              GetFtpErrorDescription = "An asynchronous request could not be made because a callback function has not been set."
    Case ERROR_INTERNET_REQUEST_PENDING:          GetFtpErrorDescription = "The required operation could not be completed because one or more requests are pending."
    Case ERROR_INTERNET_INCORRECT_FORMAT:         GetFtpErrorDescription = "The format of the request is invalid."
    Case ERROR_INTERNET_ITEM_NOT_FOUND:           GetFtpErrorDescription = "The requested item could not be located."
    Case ERROR_INTERNET_CANNOT_CONNECT:           GetFtpErrorDescription = "The attempt to connect to the server failed."
    Case ERROR_INTERNET_CONNECTION_ABORTED:       GetFtpErrorDescription = "The connection with the server has been terminated."
    Case ERROR_INTERNET_CONNECTION_RESET:         GetFtpErrorDescription = "The connection with the server has been reset."
    Case ERROR_INTERNET_FORCE_RETRY:              GetFtpErrorDescription = "The Win32 Internet function needs to redo the request."
    Case ERROR_INTERNET_INVALID_PROXY_REQUEST:    GetFtpErrorDescription = "The request to the proxy was invalid."
    Case ERROR_INTERNET_NEED_UI:                  GetFtpErrorDescription = "A user interface or other blocking operation has been requested."
    Case ERROR_INTERNET_HANDLE_EXISTS:            GetFtpErrorDescription = "The request failed because the handle already exists."
    Case ERROR_INTERNET_SEC_CERT_DATE_INVALID:    GetFtpErrorDescription = "SSL certificate date that was received from the server is bad. The certificate is expired."
    Case ERROR_INTERNET_SEC_CERT_CN_INVALID:      GetFtpErrorDescription = "SSL certificate common name (host name field) is incorrect—for example, if you entered www.server.com and the common name on the certificate says www.different.com."
    Case ERROR_INTERNET_HTTP_TO_HTTPS_ON_REDIR:   GetFtpErrorDescription = "The application is moving from a non-SSL to an SSL connection because of a redirect."
    Case ERROR_INTERNET_HTTPS_TO_HTTP_ON_REDIR:   GetFtpErrorDescription = "The application is moving from an SSL to an non-SSL connection because of a redirect."
    Case ERROR_INTERNET_MIXED_SECURITY:           GetFtpErrorDescription = "The content is not entirely secure. Some of the content being viewed may have come from unsecured servers."
    Case ERROR_INTERNET_CHG_POST_IS_NON_SECURE:   GetFtpErrorDescription = "The application is posting and attempting to change multiple lines of text on a server that is not secure."
    Case ERROR_INTERNET_POST_IS_NON_SECURE:       GetFtpErrorDescription = "The application is posting data to a sever that is not secure."
    Case ERROR_INTERNET_CLIENT_AUTH_CERT_NEEDED:  GetFtpErrorDescription = "The server is requesting client authentication."
    Case ERROR_INTERNET_INVALID_CA:               GetFtpErrorDescription = "The function is unfamiliar with the Certificate Authority that generated the server's certificate."
    Case ERROR_INTERNET_CLIENT_AUTH_NOT_SETUP:    GetFtpErrorDescription = "Client authorization is not set up on this computer."
    Case ERROR_INTERNET_ASYNC_THREAD_FAILED:      GetFtpErrorDescription = "The application could not start an asynchronous thread."
    Case ERROR_INTERNET_REDIRECT_SCHEME_CHANGE:   GetFtpErrorDescription = "The function could not handle the redirection, because the scheme changed (for example, HTTP to FTP)."
    Case ERROR_INTERNET_DIALOG_PENDING:           GetFtpErrorDescription = "Another thread has a password dialog box in progress."
    Case ERROR_INTERNET_RETRY_DIALOG:             GetFtpErrorDescription = "The dialog box should be retried."
    Case ERROR_INTERNET_HTTPS_HTTP_SUBMIT_REDIR:  GetFtpErrorDescription = "The data being submitted to an SSL connection is being redirected to a non-SSL connection."
    Case ERROR_INTERNET_INSERT_CDROM:             GetFtpErrorDescription = "The request requires a CD-ROM to be inserted in the CD-ROM drive to locate the resource requested."
    Case ERROR_INTERNET_FORTEZZA_LOGIN_NEEDED:    GetFtpErrorDescription = "The requested resource requires Fortezza authentication."
    Case ERROR_INTERNET_SEC_CERT_ERRORS:          GetFtpErrorDescription = "The SSL certificate contains errors."
    Case ERROR_INTERNET_SEC_CERT_NO_REV:          GetFtpErrorDescription = "SSL certificate had no REV"
    Case ERROR_INTERNET_SEC_CERT_REV_FAILED:      GetFtpErrorDescription = "SSL certificate REV failed"
    
    ' FTP API Errors
    Case ERROR_FTP_TRANSFER_IN_PROGRESS:          GetFtpErrorDescription = "The requested operation cannot be made on the FTP session handle because an operation is already in progress."
    Case ERROR_FTP_DROPPED:                       GetFtpErrorDescription = "The FTP operation was not completed because the session was aborted."
    Case ERROR_FTP_NO_PASSIVE_MODE:               GetFtpErrorDescription = "Passive mode is not available on the server."
    
    ' Gopher API Errors
    Case ERROR_GOPHER_PROTOCOL_ERROR:             GetFtpErrorDescription = "An error was detected while parsing data returned from the Gopher server."
    Case ERROR_GOPHER_NOT_FILE:                   GetFtpErrorDescription = "The request must be made for a file locator."
    Case ERROR_GOPHER_DATA_ERROR:                 GetFtpErrorDescription = "An error was detected while receiving data from the Gopher server."
    Case ERROR_GOPHER_END_OF_DATA:                GetFtpErrorDescription = "The end of the data has been reached."
    Case ERROR_GOPHER_INVALID_LOCATOR:            GetFtpErrorDescription = "The supplied locator is not valid."
    Case ERROR_GOPHER_INCORRECT_LOCATOR_TYPE:     GetFtpErrorDescription = "The type of the locator is not correct for this operation."
    Case ERROR_GOPHER_NOT_GOPHER_PLUS:            GetFtpErrorDescription = "The requested operation can be made only against a Gopher+ server, or with a locator that specifies a Gopher+ operation."
    Case ERROR_GOPHER_ATTRIBUTE_NOT_FOUND:        GetFtpErrorDescription = "The requested attribute could not be located."
    Case ERROR_GOPHER_UNKNOWN_LOCATOR:            GetFtpErrorDescription = "The locator type is unknown."
    
    ' HTTP API Errors
    Case ERROR_HTTP_HEADER_NOT_FOUND:             GetFtpErrorDescription = "The requested header could not be located."
    Case ERROR_HTTP_DOWNLEVEL_SERVER:             GetFtpErrorDescription = "The server did not return any headers."
    Case ERROR_HTTP_INVALID_SERVER_RESPONSE:      GetFtpErrorDescription = "The server response could not be parsed."
    Case ERROR_HTTP_INVALID_HEADER:               GetFtpErrorDescription = "The supplied header is invalid."
    Case ERROR_HTTP_INVALID_QUERY_REQUEST:        GetFtpErrorDescription = "The request made to HttpQueryInfo is invalid."
    Case ERROR_HTTP_HEADER_ALREADY_EXISTS:        GetFtpErrorDescription = "The header could not be added because it already exists."
    Case ERROR_HTTP_REDIRECT_FAILED:              GetFtpErrorDescription = "The redirection failed because either the scheme changed (for example, HTTP to FTP) or all attempts made to redirect failed (default is five attempts)."
    Case ERROR_HTTP_NOT_REDIRECTED:               GetFtpErrorDescription = "The HTTP request was not redirected."
    Case ERROR_HTTP_COOKIE_NEEDS_CONFIRMATION:    GetFtpErrorDescription = "The HTTP cookie requires confirmation."
    Case ERROR_HTTP_COOKIE_DECLINED:              GetFtpErrorDescription = "The HTTP cookie was declined by the server."
    Case ERROR_HTTP_REDIRECT_NEEDS_CONFIRMATION:  GetFtpErrorDescription = "The redirection requires user confirmation."
    
    ' Additional Internet API Error Codes
    Case ERROR_INTERNET_SECURITY_CHANNEL_ERROR:   GetFtpErrorDescription = "'The application experienced an internal error loading the SSL libraries."
    Case ERROR_INTERNET_UNABLE_TO_CACHE_FILE:     GetFtpErrorDescription = "The function was unable to cache the file."
    Case ERROR_INTERNET_TCPIP_NOT_INSTALLED:      GetFtpErrorDescription = "The required protocol stack is not loaded and the application cannot start WinSock."
    Case ERROR_INTERNET_DISCONNECTED:             GetFtpErrorDescription = "The Internet connection has been lost."
    Case ERROR_INTERNET_SERVER_UNREACHABLE:       GetFtpErrorDescription = "The Web site or server indicated is unreachable."
    Case ERROR_INTERNET_PROXY_SERVER_UNREACHABLE: GetFtpErrorDescription = "The designated proxy server cannot be reached."
    Case ERROR_INTERNET_BAD_AUTO_PROXY_SCRIPT:    GetFtpErrorDescription = "There was an error in the automatic proxy configuration script."
    Case ERROR_INTERNET_UNABLE_TO_DOWNLOAD_SCRIPT: GetFtpErrorDescription = "The automatic proxy configuration script could not be downloaded. The INTERNET_FLAG_MUST_CACHE_REQUEST flag was set."
    Case ERROR_INTERNET_SEC_INVALID_CERT:         GetFtpErrorDescription = "SSL certificate is invalid."
    Case ERROR_INTERNET_SEC_CERT_REVOKED:         GetFtpErrorDescription = "SSL certificate was revoked."
    
    ' InternetAutodial Specific Errors
    Case ERROR_INTERNET_FAILED_DUETOSECURITYCHECK: GetFtpErrorDescription = "The function failed due to a security check."
    Case ERROR_INTERNET_NOT_INITIALIZED:          GetFtpErrorDescription = "Initialization of the Win32 Internet API has not occurred. Indicates that a higher-level function, such as InternetOpen, has not been called yet."
    Case ERROR_INTERNET_NEED_MSN_SSPI_PKG:        GetFtpErrorDescription = "Not currently implemented."
    Case ERROR_INTERNET_LOGIN_FAILURE_DISPLAY_ENTITY_BODY: GetFtpErrorDescription = "The MS-Logoff digest header has been returned from the Web site. This header specifically instructs the digest package to purge credentials for the associated realm. This error will only be returned if INTERNET_ERROR_MASK_LOGIN_FAILURE_DISPLAY_ENTITY_BODY has been set."
  End Select
  
End Function
