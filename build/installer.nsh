; Math Popup — NSIS installer customizations (auto-included by electron-builder).
;
; Force-close any running Math Popup (and its Electron helper processes) before
; the installer removes the previously installed version. Without this, an update
; or reinstall can fail with NSIS error 2 — "Failed to uninstall old application
; files" — if a lingering process still holds a handle to a file in the old
; install folder (the most likely cause of that error during auto-updates).
;
; customInit runs early in the installer's .onInit, before the old version is
; uninstalled. The taskkill exit code is ignored: a non-zero code just means the
; app was not running, which is fine.
!macro customInit
  Push $0
  ; /F = force, /T = also kill child processes (GPU / renderer / crashpad).
  nsExec::Exec 'taskkill /F /T /IM "Math Popup.exe"'
  Pop $0
  ; Give Windows a moment to release the file handles before the uninstall step.
  Sleep 1500
  Pop $0
!macroend
