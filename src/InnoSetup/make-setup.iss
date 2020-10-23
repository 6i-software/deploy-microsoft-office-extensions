; -------------------------------------------------------------------------------------
; @copyright 6i (2020)
; @author 20100 <vb20100bv@gmail.com>
; Released under a MIT license.
;
; This InnoSetup script can install Microsoft Office extensions & addins, and can
; activate Excel Addins wich is stored in *.xlam file. This script is inspired to
; work of Bovender.
;
; Thanks to work of Daniel's XL Toolbox (xltoolbox.sf.net - http://github.com/bovender)
; -------------------------------------------------------------------------------------

; Constants
#define AppName "MicrosoftOfficeExtensions"
#define AppVersion "1.0.0"
#define Year "2020"
#define AppExeName AppName + "_v" + AppVersion + ".exe"
#define AppURL "http://www.example.com/"
#define AppCorporate "6i"
#define AppCopyright "Copyright (C) "+ Year + ", " + AppCorporate
#define AppContact "email@acme.com"

; Assets directory use to build setup
#define AssetsDir SourcePath + "assets"
; Output directory where setup release is build
#define OutputDir SourcePath + "..\..\release\v" + AppVersion + "\"
; Source directory where Macro VBA Office Excel and Word are stored
#define SourceDir SourcePath + "..\"
; Name of log file used to debug installation. Required activation of SetupLogging=True
#define LogFile AppName + "_v" + AppVersion + "_setup.log"

; The value of AppId uniquely identifies this application. Do not use the same AppId value
; in installers for other applications. To generate a new GUID inside the InnoSetup IDE,
; click Tools and choose Generate GUID.
#define AppGUID "BE632A57-6B04-4A39-A97D-10EF58B76B3A"
#define AppID "6i_MOExtensions_" + AppGUID

[Setup]
AppId={#AppID}
AppName={#AppName}
AppVersion={#AppVersion}
AppVerName={#AppName} version {#AppVersion}
AppPublisher={#AppCorporate}
AppPublisherURL={#AppURL}
AppSupportURL={#AppURL}
AppUpdatesURL={#AppURL}
AppCopyright={#AppCopyright}
AppContact={#AppContact}

LicenseFile={#AssetsDir}\LICENCE.txt
InfoBeforeFile={#AssetsDir}\BEFORE_INSTALL.txt
InfoAfterFile={#AssetsDir}\AFTER_INSTALL.txt

VersionInfoVersion={#AppVersion}
VersionInfoCompany={#AppCorporate}
VersionInfoProductName={#AppName}
VersionInfoCopyright={#AppCopyright}
VersionInfoDescription=Installeur {#AppName} {#AppVersion}
VersionInfoProductVersion={#AppVersion}
VersionInfoTextVersion={#AppVersion}

DefaultGroupName={#AppCorporate}\{#AppName}
DisableProgramGroupPage=false
DisableDirPage=true
DisableWelcomePage=false
CreateAppDir=true
AppendDefaultDirName=false

; Setup output options
OutputDir={#OutputDir}
OutputBaseFilename={#AppCorporate}_{#AppName}_v{#AppVersion}
Compression=lzma2/fast
SolidCompression=yes
SetupIconFile={#AssetsDir}\setup.ico

; The destination folder is also determined with code section, since different language versions of Excel expect
; addins in localized folders.
DefaultDirName={userappdata}\Microsoft\AddIns\

; The uninstall icon must be included in the setup package and placed in the output folder.
UninstallDisplayIcon={app}\setup.ico

; Make this setup program work with 32-bit and 64-bit Windows
ArchitecturesAllowed=x86 x64
ArchitecturesInstallIn64BitMode=x64

; Always write a log file (set false/true to disable/enable logging feature)
SetupLogging=true

; Allow normal users to install the addin into their profile. This directive also ensures that the
; uninstall information is stored in the user profile rather than a system folder, which would
; require administrative rights. To run in non administrative install mode, i.e. install for current user only
; we must set PrivilegesRequired=lowest
PrivilegesRequired=admin
; PrivilegesRequiredOverridesAllowed=dialog

; Section to code signing setup with a certificate
; SignTool=signtool sign /v /s 6i-ACME /n $q6i ACME$q /t http://timestamp.verisign.com/scripts/timstamp.dll /d $qCertification MicrosoftOfficeExtensions$q $f

; Setup wizard options
WizardImageFile={#AssetsDir}\images\innosetup_background.bmp
WizardImageStretch=no
WizardSmallImageFile={#AssetsDir}\images\logo.bmp
BackColor=$FFFF00

[Languages]
#include "include/Languages.iss"

[CustomMessages]
#include "include/CustomMessages.iss"

[Files]
#include "include/Files.iss"

[Tasks]
#include "include/Tasks.iss"

[Icons]
#include "include/Icons.iss"

[Run]
#include "include/Run.iss"

[Code]
#include "include/Code.iss"