; -------------------------------------------------------------------------------------
; @copyright 6i (2020)
; @author 20100 <vb20100bv@gmail.com>
; Released under a MIT license.
; -------------------------------------------------------------------------------------

; The include file makes adds all *.xla, *.xlam files contained in extensions folder.
; All files are expected in the {#SourceDir} folder and the destination folder is
; automatically determined. It resides in the user profile.

; Include custom files in the setup project.
Source: {#AssetsDir}\setup.ico; DestDir: {code:GetDestDir}\
Source: {#SourceDir}\..\README.md; DestDir: {code:GetDestDir}\
Source: {#SourceDir}\..\LICENSE.md; DestDir: {code:GetDestDir}\
Source: {#AssetsDir}\Microsoft-Office-Extensions-example.xlsx; DestDir: {code:GetDestDir}\
Source: {#AssetsDir}\Microsoft-Office-Extensions-example.docx; DestDir: {code:GetDestDir}\

; Include all Excel addins files in the setup project
Source: {#SourceDir}\extensions\extensionsExcel\*.xlam; DestDir: {code:GetDestDir}\; Check: ShouldInstallFile(12,16); AfterInstall: ActivateAddin(12,16)
;Source: {#SourceDir}\*.xla; DestDir: {code:GetDestDir}\; Check: ShouldInstallFile(9,11); AfterInstall: ActivateAddin(9,11); Excludes: *.xlam

; Include all Word extensions files in the setup project
Source: {#SourceDir}\extensions\extensionsWord\*.dotm; DestDir: {code:GetDestDir}\..\Word\STARTUP\;
