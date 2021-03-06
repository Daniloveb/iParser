USE [iBase]
GO
/****** Object:  Table [dbo].[_Types]    Script Date: 05/10/2012 14:38:41 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
SET ANSI_PADDING ON
GO
CREATE TABLE [dbo].[_Types](
	[UID] [uniqueidentifier] ROWGUIDCOL  NOT NULL,
	[TypeDevice] [varchar](50) NOT NULL,
	[Prefix] [nchar](10) NULL,
 CONSTRAINT [PK_Types] PRIMARY KEY CLUSTERED 
(
	[UID] ASC
)WITH (PAD_INDEX  = OFF, STATISTICS_NORECOMPUTE  = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS  = ON, ALLOW_PAGE_LOCKS  = ON) ON [PRIMARY]
) ON [PRIMARY]
GO
SET ANSI_PADDING OFF
GO
/****** Object:  Table [dbo].[_iNumbers]    Script Date: 05/10/2012 14:38:41 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[_iNumbers](
	[UID] [uniqueidentifier] ROWGUIDCOL  NOT NULL,
	[INumber] [int] NOT NULL,
	[TypeID] [uniqueidentifier] NOT NULL,
 CONSTRAINT [PK_iNumbers] PRIMARY KEY CLUSTERED 
(
	[UID] ASC
)WITH (PAD_INDEX  = OFF, STATISTICS_NORECOMPUTE  = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS  = ON, ALLOW_PAGE_LOCKS  = ON) ON [PRIMARY]
) ON [PRIMARY]
GO
/****** Object:  Table [dbo].[InstalledSoft]    Script Date: 05/10/2012 14:38:41 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
SET ANSI_PADDING ON
GO
CREATE TABLE [dbo].[InstalledSoft](
	[UID] [uniqueidentifier] ROWGUIDCOL  NOT NULL,
	[INVNumberID] [uniqueidentifier] NOT NULL,
	[Date] [datetime] NULL,
	[DisplayName] [varchar](max) NULL,
	[InstallDate] [date] NULL,
	[InstallLocation] [varchar](max) NULL,
	[Puslisher] [varchar](max) NULL,
	[DisplayVersion] [varchar](max) NULL,
 CONSTRAINT [IDInstalledSoft] PRIMARY KEY CLUSTERED 
(
	[UID] ASC
)WITH (PAD_INDEX  = OFF, STATISTICS_NORECOMPUTE  = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS  = ON, ALLOW_PAGE_LOCKS  = ON) ON [PRIMARY]
) ON [PRIMARY]
GO
SET ANSI_PADDING OFF
GO
/****** Object:  Table [dbo].[Win32_VideoController]    Script Date: 05/10/2012 14:38:41 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
SET ANSI_PADDING ON
GO
CREATE TABLE [dbo].[Win32_VideoController](
	[UID] [uniqueidentifier] ROWGUIDCOL  NOT NULL,
	[INVNumberID] [uniqueidentifier] NOT NULL,
	[Date] [datetime] NULL,
	[AdapterCompatibility] [varchar](max) NULL,
	[AdapterDACType] [varchar](max) NULL,
	[AdapterRAM] [float] NULL,
	[Availability] [float] NULL,
	[Caption] [varchar](max) NULL,
	[ConfigManagerErrorCode] [float] NULL,
	[ConfigManagerUserConfig] [bit] NULL,
	[CreationClassName] [varchar](max) NULL,
	[CurrentBitsPerPixel] [float] NULL,
	[CurrentHorizontalResolution] [float] NULL,
	[CurrentNumberOfColors] [float] NULL,
	[CurrentNumberOfColumns] [float] NULL,
	[CurrentNumberOfRows] [float] NULL,
	[CurrentRefreshRate] [float] NULL,
	[CurrentScanMode] [float] NULL,
	[CurrentVerticalResolution] [float] NULL,
	[Description] [varchar](max) NULL,
	[DeviceID] [varchar](max) NULL,
	[DeviceSpecificPens] [float] NULL,
	[DriverDate] [date] NULL,
	[DriverVersion] [varchar](max) NULL,
	[InfFilename] [varchar](max) NULL,
	[InfSection] [varchar](max) NULL,
	[InstalledDisplayDrivers] [varchar](max) NULL,
	[MaxRefreshRate] [float] NULL,
	[MinRefreshRate] [float] NULL,
	[Monochrome] [bit] NULL,
	[Name] [varchar](max) NULL,
	[NumberOfColorPlanes] [float] NULL,
	[PNPDeviceID] [varchar](max) NULL,
	[Status] [varchar](max) NULL,
	[SystemCreationClassName] [varchar](max) NULL,
	[SystemName] [varchar](max) NULL,
	[VideoArchitecture] [float] NULL,
	[VideoMemoryType] [float] NULL,
	[VideoModeDescription] [varchar](max) NULL,
	[VideoProcessor] [varchar](max) NULL,
 CONSTRAINT [IDWin32_VideoController] PRIMARY KEY CLUSTERED 
(
	[UID] ASC
)WITH (PAD_INDEX  = OFF, STATISTICS_NORECOMPUTE  = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS  = ON, ALLOW_PAGE_LOCKS  = ON) ON [PRIMARY]
) ON [PRIMARY]
GO
SET ANSI_PADDING OFF
GO
/****** Object:  Table [dbo].[Win32_Service]    Script Date: 05/10/2012 14:38:41 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
SET ANSI_PADDING ON
GO
CREATE TABLE [dbo].[Win32_Service](
	[UID] [uniqueidentifier] ROWGUIDCOL  NOT NULL,
	[INVNumberID] [uniqueidentifier] NOT NULL,
	[Date] [datetime] NULL,
	[AcceptPause] [bit] NULL,
	[AcceptStop] [bit] NULL,
	[Caption] [varchar](max) NULL,
	[CheckPoint] [float] NULL,
	[CreationClassName] [varchar](max) NULL,
	[Description] [varchar](max) NULL,
	[DesktopInteract] [bit] NULL,
	[DisplayName] [varchar](max) NULL,
	[ErrorControl] [varchar](max) NULL,
	[ExitCode] [float] NULL,
	[Name] [varchar](max) NULL,
	[PathName] [varchar](max) NULL,
	[ProcessId] [float] NULL,
	[ServiceSpecificExitCode] [float] NULL,
	[ServiceType] [varchar](max) NULL,
	[Started] [bit] NULL,
	[StartMode] [varchar](max) NULL,
	[StartName] [varchar](max) NULL,
	[State] [varchar](max) NULL,
	[Status] [varchar](max) NULL,
	[SystemCreationClassName] [varchar](max) NULL,
	[SystemName] [varchar](max) NULL,
	[TagId] [float] NULL,
	[WaitHint] [float] NULL,
 CONSTRAINT [IDWin32_Service] PRIMARY KEY CLUSTERED 
(
	[UID] ASC
)WITH (PAD_INDEX  = OFF, STATISTICS_NORECOMPUTE  = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS  = ON, ALLOW_PAGE_LOCKS  = ON) ON [PRIMARY]
) ON [PRIMARY]
GO
SET ANSI_PADDING OFF
GO
/****** Object:  Table [dbo].[Win32_QuickFixEngineering]    Script Date: 05/10/2012 14:38:41 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
SET ANSI_PADDING ON
GO
CREATE TABLE [dbo].[Win32_QuickFixEngineering](
	[UID] [uniqueidentifier] ROWGUIDCOL  NOT NULL,
	[INVNumberID] [uniqueidentifier] NOT NULL,
	[Date] [datetime] NULL,
	[CSName] [varchar](max) NULL,
	[HotFixID] [varchar](max) NULL,
	[ServicePackInEffect] [varchar](max) NULL,
	[Description] [varchar](max) NULL,
	[FixComments] [varchar](max) NULL,
	[InstalledBy] [varchar](max) NULL,
	[InstalledOn] [date] NULL,
 CONSTRAINT [IDWin32_QuickFixEngineering] PRIMARY KEY CLUSTERED 
(
	[UID] ASC
)WITH (PAD_INDEX  = OFF, STATISTICS_NORECOMPUTE  = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS  = ON, ALLOW_PAGE_LOCKS  = ON) ON [PRIMARY]
) ON [PRIMARY]
GO
SET ANSI_PADDING OFF
GO
/****** Object:  Table [dbo].[Win32_OperatingSystem]    Script Date: 05/10/2012 14:38:41 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
SET ANSI_PADDING ON
GO
CREATE TABLE [dbo].[Win32_OperatingSystem](
	[UID] [uniqueidentifier] ROWGUIDCOL  NOT NULL,
	[INVNumberID] [uniqueidentifier] NOT NULL,
	[Date] [datetime] NULL,
	[BootDevice] [varchar](max) NULL,
	[BuildNumber] [float] NULL,
	[BuildType] [varchar](max) NULL,
	[Caption] [varchar](max) NULL,
	[CodeSet] [float] NULL,
	[CountryCode] [float] NULL,
	[CreationClassName] [varchar](max) NULL,
	[CSCreationClassName] [varchar](max) NULL,
	[CSDVersion] [varchar](max) NULL,
	[CSName] [varchar](max) NULL,
	[CurrentTimeZone] [float] NULL,
	[DataExecutionPrevention_32BitApplications] [bit] NULL,
	[DataExecutionPrevention_Available] [bit] NULL,
	[DataExecutionPrevention_Drivers] [bit] NULL,
	[DataExecutionPrevention_SupportPolicy] [float] NULL,
	[Debug] [bit] NULL,
	[Distributed] [bit] NULL,
	[EncryptionLevel] [float] NULL,
	[ForegroundApplicationBoost] [float] NULL,
	[FreePhysicalMemory] [float] NULL,
	[FreeSpaceInPagingFiles] [float] NULL,
	[FreeVirtualMemory] [float] NULL,
	[InstallDate] [date] NULL,
	[LargeSystemCache] [float] NULL,
	[LastBootUpTime] [date] NULL,
	[LocalDateTime] [date] NULL,
	[Locale] [float] NULL,
	[Manufacturer] [varchar](max) NULL,
	[MaxNumberOfProcesses] [float] NULL,
	[MaxProcessMemorySize] [float] NULL,
	[Name] [varchar](max) NULL,
	[NumberOfProcesses] [float] NULL,
	[NumberOfUsers] [float] NULL,
	[Organization] [float] NULL,
	[OSLanguage] [float] NULL,
	[OSType] [float] NULL,
	[Primary] [bit] NULL,
	[ProductType] [float] NULL,
	[QuantumLength] [float] NULL,
	[QuantumType] [float] NULL,
	[RegisteredUser] [varchar](max) NULL,
	[SerialNumber] [varchar](max) NULL,
	[ServicePackMajorVersion] [float] NULL,
	[ServicePackMinorVersion] [float] NULL,
	[SizeStoredInPagingFiles] [float] NULL,
	[Status] [varchar](max) NULL,
	[SuiteMask] [float] NULL,
	[SystemDevice] [varchar](max) NULL,
	[SystemDirectory] [varchar](max) NULL,
	[SystemDrive] [varchar](max) NULL,
	[TotalVirtualMemorySize] [float] NULL,
	[TotalVisibleMemorySize] [float] NULL,
	[Version] [date] NULL,
	[WindowsDirectory] [varchar](max) NULL,
 CONSTRAINT [IDWin32_OperatingSystem] PRIMARY KEY CLUSTERED 
(
	[UID] ASC
)WITH (PAD_INDEX  = OFF, STATISTICS_NORECOMPUTE  = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS  = ON, ALLOW_PAGE_LOCKS  = ON) ON [PRIMARY]
) ON [PRIMARY]
GO
SET ANSI_PADDING OFF
GO
/****** Object:  Table [dbo].[Win32_NetworkAdapterConfiguration]    Script Date: 05/10/2012 14:38:41 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
SET ANSI_PADDING ON
GO
CREATE TABLE [dbo].[Win32_NetworkAdapterConfiguration](
	[UID] [uniqueidentifier] ROWGUIDCOL  NOT NULL,
	[INVNumberID] [uniqueidentifier] NOT NULL,
	[Date] [datetime] NULL,
	[Caption] [varchar](max) NULL,
	[Description] [varchar](max) NULL,
	[DHCPEnabled] [bit] NULL,
	[Index] [float] NULL,
	[IPEnabled] [bit] NULL,
	[IPXEnabled] [bit] NULL,
	[ServiceName] [varchar](max) NULL,
	[SettingID] [varchar](max) NULL,
	[MACAddress] [varchar](max) NULL,
	[DatabasePath] [varchar](max) NULL,
	[DNSEnabledForWINSResolution] [bit] NULL,
	[DNSHostName] [varchar](max) NULL,
	[DNSServerSearchOrder] [varchar](max) NULL,
	[DomainDNSRegistrationEnabled] [bit] NULL,
	[FullDNSRegistrationEnabled] [bit] NULL,
	[IPAddress] [varchar](max) NULL,
	[IPConnectionMetric] [float] NULL,
	[IPFilterSecurityEnabled] [bit] NULL,
	[IPSecPermitIPProtocols] [float] NULL,
	[IPSecPermitTCPPorts] [float] NULL,
	[IPSecPermitUDPPorts] [float] NULL,
	[TcpipNetbiosOptions] [float] NULL,
	[WINSEnableLMHostsLookup] [bit] NULL,
 CONSTRAINT [IDWin32_NetworkAdapterConfiguration] PRIMARY KEY CLUSTERED 
(
	[UID] ASC
)WITH (PAD_INDEX  = OFF, STATISTICS_NORECOMPUTE  = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS  = ON, ALLOW_PAGE_LOCKS  = ON) ON [PRIMARY]
) ON [PRIMARY]
GO
SET ANSI_PADDING OFF
GO
/****** Object:  Table [dbo].[Win32_NetworkAdapter]    Script Date: 05/10/2012 14:38:41 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
SET ANSI_PADDING ON
GO
CREATE TABLE [dbo].[Win32_NetworkAdapter](
	[UID] [uniqueidentifier] ROWGUIDCOL  NOT NULL,
	[INVNumberID] [uniqueidentifier] NOT NULL,
	[Date] [datetime] NULL,
	[Availability] [float] NULL,
	[Caption] [varchar](max) NULL,
	[CreationClassName] [varchar](max) NULL,
	[Description] [varchar](max) NULL,
	[DeviceID] [float] NULL,
	[Index] [float] NULL,
	[Installed] [bit] NULL,
	[MaxNumberControlled] [float] NULL,
	[Name] [varchar](max) NULL,
	[PowerManagementSupported] [bit] NULL,
	[ProductName] [varchar](max) NULL,
	[SystemCreationClassName] [varchar](max) NULL,
	[SystemName] [varchar](max) NULL,
	[TimeOfLastReset] [date] NULL,
	[ConfigManagerErrorCode] [float] NULL,
	[ConfigManagerUserConfig] [bit] NULL,
	[Manufacturer] [varchar](max) NULL,
	[PNPDeviceID] [varchar](max) NULL,
	[ServiceName] [varchar](max) NULL,
	[AdapterType] [varchar](max) NULL,
	[AdapterTypeId] [float] NULL,
	[MACAddress] [varchar](max) NULL,
	[NetConnectionID] [varchar](max) NULL,
	[NetConnectionStatus] [float] NULL,
 CONSTRAINT [IDWin32_NetworkAdapter] PRIMARY KEY CLUSTERED 
(
	[UID] ASC
)WITH (PAD_INDEX  = OFF, STATISTICS_NORECOMPUTE  = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS  = ON, ALLOW_PAGE_LOCKS  = ON) ON [PRIMARY]
) ON [PRIMARY]
GO
SET ANSI_PADDING OFF
GO
/****** Object:  Table [dbo].[Win32_MemoryDevice]    Script Date: 05/10/2012 14:38:41 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
SET ANSI_PADDING ON
GO
CREATE TABLE [dbo].[Win32_MemoryDevice](
	[UID] [uniqueidentifier] ROWGUIDCOL  NOT NULL,
	[INVNumberID] [uniqueidentifier] NOT NULL,
	[Date] [datetime] NULL,
	[Caption] [varchar](max) NULL,
	[CreationClassName] [varchar](max) NULL,
	[Description] [varchar](max) NULL,
	[DeviceID] [varchar](max) NULL,
	[EndingAddress] [float] NULL,
	[Name] [varchar](max) NULL,
	[StartingAddress] [float] NULL,
	[SystemCreationClassName] [varchar](max) NULL,
	[SystemName] [varchar](max) NULL,
 CONSTRAINT [IDWin32_MemoryDevice] PRIMARY KEY CLUSTERED 
(
	[UID] ASC
)WITH (PAD_INDEX  = OFF, STATISTICS_NORECOMPUTE  = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS  = ON, ALLOW_PAGE_LOCKS  = ON) ON [PRIMARY]
) ON [PRIMARY]
GO
SET ANSI_PADDING OFF
GO
/****** Object:  Table [dbo].[Win32_LogicalDisk]    Script Date: 05/10/2012 14:38:41 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
SET ANSI_PADDING ON
GO
CREATE TABLE [dbo].[Win32_LogicalDisk](
	[UID] [uniqueidentifier] ROWGUIDCOL  NOT NULL,
	[INVNumberID] [uniqueidentifier] NOT NULL,
	[Date] [datetime] NULL,
	[Caption] [varchar](max) NULL,
	[CreationClassName] [varchar](max) NULL,
	[Description] [varchar](max) NULL,
	[DeviceID] [varchar](max) NULL,
	[DriveType] [float] NULL,
	[MediaType] [float] NULL,
	[Name] [varchar](max) NULL,
	[SystemCreationClassName] [varchar](max) NULL,
	[SystemName] [varchar](max) NULL,
	[Compressed] [bit] NULL,
	[FileSystem] [varchar](max) NULL,
	[FreeSpace] [float] NULL,
	[MaximumComponentLength] [float] NULL,
	[QuotasDisabled] [bit] NULL,
	[QuotasIncomplete] [bit] NULL,
	[QuotasRebuilding] [bit] NULL,
	[Size] [float] NULL,
	[SupportsDiskQuotas] [bit] NULL,
	[SupportsFileBasedCompression] [bit] NULL,
	[VolumeDirty] [bit] NULL,
	[VolumeSerialNumber] [varchar](max) NULL,
	[VolumeName] [varchar](max) NULL,
 CONSTRAINT [IDWin32_LogicalDisk] PRIMARY KEY CLUSTERED 
(
	[UID] ASC
)WITH (PAD_INDEX  = OFF, STATISTICS_NORECOMPUTE  = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS  = ON, ALLOW_PAGE_LOCKS  = ON) ON [PRIMARY]
) ON [PRIMARY]
GO
SET ANSI_PADDING OFF
GO
/****** Object:  Table [dbo].[Win32_DiskDrive]    Script Date: 05/10/2012 14:38:41 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
SET ANSI_PADDING ON
GO
CREATE TABLE [dbo].[Win32_DiskDrive](
	[UID] [uniqueidentifier] ROWGUIDCOL  NOT NULL,
	[INVNumberID] [uniqueidentifier] NOT NULL,
	[Date] [datetime] NULL,
	[BytesPerSector] [float] NULL,
	[Capabilities] [varchar](max) NULL,
	[Caption] [varchar](max) NULL,
	[ConfigManagerErrorCode] [float] NULL,
	[ConfigManagerUserConfig] [bit] NULL,
	[CreationClassName] [varchar](max) NULL,
	[Description] [varchar](max) NULL,
	[DeviceID] [varchar](max) NULL,
	[Index] [float] NULL,
	[InterfaceType] [varchar](max) NULL,
	[Manufacturer] [varchar](max) NULL,
	[MediaLoaded] [bit] NULL,
	[MediaType] [varchar](max) NULL,
	[Model] [varchar](max) NULL,
	[Name] [varchar](max) NULL,
	[Partitions] [float] NULL,
	[PNPDeviceID] [varchar](max) NULL,
	[SCSIBus] [float] NULL,
	[SCSILogicalUnit] [float] NULL,
	[SCSIPort] [float] NULL,
	[SCSITargetId] [float] NULL,
	[SectorsPerTrack] [float] NULL,
	[Signature] [float] NULL,
	[Size] [float] NULL,
	[Status] [varchar](max) NULL,
	[SystemCreationClassName] [varchar](max) NULL,
	[SystemName] [varchar](max) NULL,
	[TotalCylinders] [float] NULL,
	[TotalHeads] [float] NULL,
	[TotalSectors] [float] NULL,
	[TotalTracks] [float] NULL,
	[TracksPerCylinder] [float] NULL,
 CONSTRAINT [IDWin32_DiskDrive] PRIMARY KEY CLUSTERED 
(
	[UID] ASC
)WITH (PAD_INDEX  = OFF, STATISTICS_NORECOMPUTE  = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS  = ON, ALLOW_PAGE_LOCKS  = ON) ON [PRIMARY]
) ON [PRIMARY]
GO
SET ANSI_PADDING OFF
GO
/****** Object:  Table [dbo].[Win32_ComputerSystem]    Script Date: 05/10/2012 14:38:41 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
SET ANSI_PADDING ON
GO
CREATE TABLE [dbo].[Win32_ComputerSystem](
	[UID] [uniqueidentifier] ROWGUIDCOL  NOT NULL,
	[INVNumberID] [uniqueidentifier] NOT NULL,
	[Date] [datetime] NULL,
	[AdminPasswordStatus] [float] NULL,
	[AutomaticResetBootOption] [bit] NULL,
	[AutomaticResetCapability] [bit] NULL,
	[BootROMSupported] [bit] NULL,
	[BootupState] [varchar](max) NULL,
	[Caption] [varchar](max) NULL,
	[ChassisBootupState] [float] NULL,
	[CreationClassName] [varchar](max) NULL,
	[CurrentTimeZone] [float] NULL,
	[DaylightInEffect] [bit] NULL,
	[Description] [varchar](max) NULL,
	[Domain] [varchar](max) NULL,
	[DomainRole] [float] NULL,
	[EnableDaylightSavingsTime] [bit] NULL,
	[FrontPanelResetStatus] [float] NULL,
	[InfraredSupported] [bit] NULL,
	[KeyboardPasswordStatus] [float] NULL,
	[Manufacturer] [varchar](max) NULL,
	[Model] [varchar](max) NULL,
	[Name] [varchar](max) NULL,
	[NetworkServerModeEnabled] [bit] NULL,
	[NumberOfLogicalProcessors] [float] NULL,
	[NumberOfProcessors] [float] NULL,
	[PartOfDomain] [bit] NULL,
	[PauseAfterReset] [float] NULL,
	[PowerOnPasswordStatus] [float] NULL,
	[PowerState] [float] NULL,
	[PowerSupplyState] [float] NULL,
	[PrimaryOwnerName] [varchar](max) NULL,
	[ResetCapability] [float] NULL,
	[ResetCount] [float] NULL,
	[ResetLimit] [float] NULL,
	[Roles] [varchar](max) NULL,
	[Status] [varchar](max) NULL,
	[SystemStartupDelay] [float] NULL,
	[SystemStartupOptions] [varchar](max) NULL,
	[SystemStartupSetting] [float] NULL,
	[SystemType] [varchar](max) NULL,
	[ThermalState] [float] NULL,
	[TotalPhysicalMemory] [float] NULL,
	[UserName] [varchar](max) NULL,
	[WakeUpType] [float] NULL,
 CONSTRAINT [IDWin32_ComputerSystem] PRIMARY KEY CLUSTERED 
(
	[UID] ASC
)WITH (PAD_INDEX  = OFF, STATISTICS_NORECOMPUTE  = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS  = ON, ALLOW_PAGE_LOCKS  = ON) ON [PRIMARY]
) ON [PRIMARY]
GO
SET ANSI_PADDING OFF
GO
/****** Object:  Table [dbo].[Win32_BIOS]    Script Date: 05/10/2012 14:38:41 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
SET ANSI_PADDING ON
GO
CREATE TABLE [dbo].[Win32_BIOS](
	[UID] [uniqueidentifier] ROWGUIDCOL  NOT NULL,
	[INVNumberID] [uniqueidentifier] NOT NULL,
	[Date] [datetime] NULL,
	[BiosCharacteristics] [varchar](max) NULL,
	[BIOSVersion] [varchar](max) NULL,
	[Caption] [varchar](max) NULL,
	[CurrentLanguage] [varchar](max) NULL,
	[Description] [varchar](max) NULL,
	[InstallableLanguages] [float] NULL,
	[ListOfLanguages] [varchar](max) NULL,
	[Manufacturer] [varchar](max) NULL,
	[Name] [varchar](max) NULL,
	[PrimaryBIOS] [bit] NULL,
	[ReleaseDate] [date] NULL,
	[SMBIOSBIOSVersion] [varchar](max) NULL,
	[SMBIOSMajorVersion] [float] NULL,
	[SMBIOSMinorVersion] [float] NULL,
	[SMBIOSPresent] [bit] NULL,
	[SoftwareElementID] [varchar](max) NULL,
	[SoftwareElementState] [float] NULL,
	[Status] [varchar](max) NULL,
	[TargetOperatingSystem] [float] NULL,
	[Version] [varchar](max) NULL,
 CONSTRAINT [IDWin32_BIOS] PRIMARY KEY CLUSTERED 
(
	[UID] ASC
)WITH (PAD_INDEX  = OFF, STATISTICS_NORECOMPUTE  = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS  = ON, ALLOW_PAGE_LOCKS  = ON) ON [PRIMARY]
) ON [PRIMARY]
GO
SET ANSI_PADDING OFF
GO
/****** Object:  Table [dbo].[Win32_BaseBoard]    Script Date: 05/10/2012 14:38:41 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
SET ANSI_PADDING ON
GO
CREATE TABLE [dbo].[Win32_BaseBoard](
	[UID] [uniqueidentifier] ROWGUIDCOL  NOT NULL,
	[INVNumberID] [uniqueidentifier] NOT NULL,
	[Date] [datetime] NULL,
	[Caption] [varchar](max) NULL,
	[CreationClassName] [varchar](max) NULL,
	[Description] [varchar](max) NULL,
	[HostingBoard] [bit] NULL,
	[Manufacturer] [varchar](max) NULL,
	[Name] [varchar](max) NULL,
	[PoweredOn] [bit] NULL,
	[Product] [varchar](max) NULL,
	[Tag] [varchar](max) NULL,
	[Version] [varchar](max) NULL,
 CONSTRAINT [IDWin32_BaseBoard] PRIMARY KEY CLUSTERED 
(
	[UID] ASC
)WITH (PAD_INDEX  = OFF, STATISTICS_NORECOMPUTE  = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS  = ON, ALLOW_PAGE_LOCKS  = ON) ON [PRIMARY]
) ON [PRIMARY]
GO
SET ANSI_PADDING OFF
GO
/****** Object:  Table [dbo].[NetworkDeviceInfo]    Script Date: 05/10/2012 14:38:41 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
SET ANSI_PADDING ON
GO
CREATE TABLE [dbo].[NetworkDeviceInfo](
	[UID] [uniqueidentifier] ROWGUIDCOL  NOT NULL,
	[INVNumberID] [uniqueidentifier] NOT NULL,
	[Date] [datetime] NULL,
	[MACAdress] [varchar](max) NULL,
	[Model] [varchar](max) NULL,
 CONSTRAINT [IDNetworkDeviceInfo] PRIMARY KEY CLUSTERED 
(
	[UID] ASC
)WITH (PAD_INDEX  = OFF, STATISTICS_NORECOMPUTE  = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS  = ON, ALLOW_PAGE_LOCKS  = ON) ON [PRIMARY]
) ON [PRIMARY]
GO
SET ANSI_PADDING OFF
GO
/****** Object:  View [dbo].[MACs]    Script Date: 05/10/2012 14:38:42 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
create View [dbo].[MACs]
as
select N.INumber,I.MACAddress from 
(select * from Win32_NetworkAdapter where NetConnectionID is not NULL and MACAddress is not NULL and ServiceName <> 'vmnetadapter' and Name <> 'сетевой адаптер 1394') I
left join
(select * from iNumbers) N
on I.INVNumberID = N.UID
union
select NN.INumber,Y.MACAdress from 
(select * from NetworkDeviceInfo) Y
left join
(select * from iNumbers) NN
on Y.INVNumberID=NN.UID
GO
/****** Object:  Default [DF_iNumbers_UID]    Script Date: 05/10/2012 14:38:41 ******/
ALTER TABLE [dbo].[_iNumbers] ADD  CONSTRAINT [DF_iNumbers_UID]  DEFAULT (newid()) FOR [UID]
GO
/****** Object:  Default [DF_Types_UID]    Script Date: 05/10/2012 14:38:41 ******/
ALTER TABLE [dbo].[_Types] ADD  CONSTRAINT [DF_Types_UID]  DEFAULT (newid()) FOR [UID]
GO
