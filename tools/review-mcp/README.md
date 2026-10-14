# Access Review MCP server

An MCP server that lets Claude review the forms of a Microsoft Access front end without using your screen. It starts its own Access instance over COM, keeps every Access window just past the edge of the desktop, and captures forms with `PrintWindow`, which asks Access to draw a window into a bitmap. Nothing comes to the front, nothing takes focus, and no mouse or keyboard input is sent, so you can keep working on the same PC while a review runs.

It was built for visual review of apps made with [access-modern-ui](../../README.md), but the core tools work with any Access front end.

```
AccessReview.ps1      the review functions (usable from plain PowerShell too)
AccessReviewMcp.ps1   the MCP server (stdio) that exposes them as tools
manifest.json         desktop extension manifest (for the .mcpb package)
```

## Requirements

- Windows 10 or 11, or Windows Server with a desktop session.
- Microsoft Access, desktop version. Tested with Access in Microsoft 365; the free Access Runtime is untested.
- Windows PowerShell 5.1 and .NET Framework 4.x, which come with Windows. Nothing else is installed.
- A front end you trust, in a trusted location. Under automation Access runs startup code (AutoExec, startup form) without the Enable Content bar.

## Install

### Claude Desktop: extension file

1. Get `access-review-<version>.mcpb` (or build it, see [Build the package](#build-the-package)).
2. In Claude Desktop open **Settings > Extensions**, then drag the file in (or use **Advanced settings > Install extension**).
3. Fill in the optional settings (see [Settings](#settings)) and turn it on.

### Claude Desktop: config file

In Claude Desktop open **Settings > Developer > Edit Config**. That opens `claude_desktop_config.json`. Add an `access-review` entry under `mcpServers`, using the folder you put the scripts in, then restart Claude Desktop:

```json
{
  "mcpServers": {
    "access-review": {
      "command": "powershell.exe",
      "args": ["-NoLogo", "-NoProfile", "-NonInteractive", "-ExecutionPolicy", "Bypass",
               "-File", "C:\\Tools\\access-review\\AccessReviewMcp.ps1"]
    }
  }
}
```

If the file already has other servers, add the `access-review` entry next to them inside the same `mcpServers` object. Backslashes in paths are doubled in JSON.

### Claude Code

Add the same entry to a project's `.mcp.json`, or run:

```
claude mcp add access-review -- powershell.exe -NoLogo -NoProfile -NonInteractive -ExecutionPolicy Bypass -File C:\Tools\access-review\AccessReviewMcp.ps1
```

`-ExecutionPolicy Bypass` applies only to this one PowerShell process. If your organisation enforces an execution policy through Group Policy, that policy still wins; ask your administrator.

## Settings

All optional. In the extension they are fields on its settings page; with a config file they go in an `env` object on the server entry.

| Setting | Environment variable | Default | What it does |
|---|---|---|---|
| Capture folder | `ACCESS_REVIEW_OUTPUT` | `%LOCALAPPDATA%\AccessReview` | Where PNGs are saved, in a new time-stamped folder per session. `open_app` can also take `output_dir`. |
| Shared lock files | `ACCESS_REVIEW_LOCKS` | none | Full paths of lock files, separated by `;`, for a PC several reviewers take turns on. `review_lock` takes all or none; with none configured it does nothing. |
| Palettes | `ACCESS_REVIEW_PALETTES` | `Navy=1;Midnight=2;Teal Navy=6` | Palettes `set_palette` offers, as `name=number`. The number is passed to the palette procedure. The first one is the default: `close_app` switches back to it. |
| Palette procedure | `ACCESS_REVIEW_PALETTE_PROCEDURE` | `SetPalette` | Public VBA `Sub Name(ByVal i As Long)` in your front end that switches palettes. |

Example `env` for a config file:

```json
"env": {
  "ACCESS_REVIEW_OUTPUT": "D:\\Reviews",
  "ACCESS_REVIEW_PALETTES": "Light=1;Dark=2",
  "ACCESS_REVIEW_PALETTE_PROCEDURE": "ApplyTheme"
}
```

## Tools

| Tool | What it does |
|---|---|
| `open_app` | Opens an existing front end in a new Access instance (never creates a file). `mode`: `offscreen` (default) or `onscreen` for a dedicated review machine. |
| `close_app` | Resets the palette to the default, quits without saving design changes and confirms the process is gone. |
| `open_form` | Opens a form and sizes it to its design size (form width, visible section heights), or to `width_px` x `height_px` at 96 dpi. Reports whether the Access window clips it. |
| `resize_form` | Design size, a given size or maximized (Access window at the main display's work area, as a real maximized window gets). A shell form that fills the Access window is sized by resizing the Access window until its inside size matches; `SizeMatches` reports whether it did. |
| `capture` | PNG of a form, or a zoomed crop, taken once two frames in a row match (so a form still drawing isn't caught mid-draw; `Stable` reports it). Returned as an image and saved in the capture folder, named form-palette-label. |
| `controls` | Each control's Left/Top/Width/Height (96 dpi px), section, caption and colours, for measuring gaps and alignment. |
| `dialogs` | Lists and captures open message boxes and popups; `close=true` closes them. Also lists dialogs seen during earlier calls. |
| `error_log` | `mark=true` before testing, then `mark=false` to list new rows in an `error_log` table and new text in the file a public `LogFilePath()` function returns. Says so when the front end has neither. |
| `compile` | Compiles and saves all modules. |
| `run` | Runs a public procedure with `Application.Run`, for example a sign-in. Whole numbers are passed as Long (Double when too large) and output arguments come back. |
| `set_palette` | Runs the palette procedure, then reopens the open forms at their last size (a form left open keeps part of the old drawing). A form with a public `Reopen` method reopens itself on its current page (read from its `PageKey`), so a signed-in user stays signed in; if that fails the call returns an error and never closes the form. Other forms are closed and reopened with their OpenArgs. |
| `review_lock` / `review_unlock` | Takes and releases the shared lock files, re-reading after 3 s to catch a race. Releases only its own. |

Every Access call runs under a guard. A window that appears during a call is parked off screen, captured and described. VBA and Office message boxes are closed so the call returns. A Yes/No box ignores Close, so the guard presses No, or the default button when there is no No. Another window (a dialog form opened modally) is closed if it is still blocking the call after 4 s, except while `open_app` starts the front end, so a splash or sign-in form stays open. A call that runs past its timeout ends Access. Each tool's result lists what the guard saw.

In off-screen mode a watcher keeps every Access window, including popups the app opens or re-centres on its own, past the desktop edge for the whole session; a newly shown popup can still appear for up to about 50 ms before it is moved.

### What your front end needs

- Nothing, for opening, sizing, capturing, controls, dialogs, compile and run.
- For `error_log`: a table named `error_log` with columns `id`, `logged_at`, `error_number`, `description`, `source`, `active_form`, `active_control`, and/or a public `LogFilePath()` function returning a text log's path. [`examples/modules/modErrors.bas`](../../examples/modules/modErrors.bas) provides both.
- For `set_palette`: a public palette procedure (see [Settings](#settings)). access-modern-ui apps have `SetPalette`.

## Typical review run

1. `review_lock` with your name, if the PC is shared.
2. `open_app` on the front end, then `error_log` with `mark=true`.
3. For each page to check: `open_form` at its designed size, `capture`, `controls` to measure spacing, `dialogs`.
4. If the app has palettes: `set_palette`, `capture` again.
5. `resize_form` small (about 1100x700) and maximized, `capture`.
6. `error_log` (new entries are findings unless explained), `close_app`, `review_unlock`.

## Limits

- Hover effects, cursor changes and flicker during a real click can't be seen without a real pointer. Check those by hand.
- `PrintWindow` can't draw a minimized window, and a locked or disconnected Windows session draws nothing. The capture warns when an image comes back one solid colour. On a server, keep the review session unlocked: use auto-logon, or end Remote Desktop with `tscon <session id> /dest:console` instead of disconnecting.
- A window can't be bigger than the desktop, so the Access window is at most the desktop size (default 1920x1200). On a server with a small default resolution, raise it, or forms larger than the Access window will be clipped (`ClippedByAccessWindow` says so).
- Typing into fields and clicking buttons on a form isn't included yet. Use `run` for public procedures in the meantime.

PowerShell can't read Access form, control and section properties directly (their type information fails, so `$form.Hwnd` comes back empty). The tools read them with `Application.Eval("Forms![name]...")` and set the form's inside size late-bound.

## Using it from PowerShell

```powershell
. .\AccessReview.ps1
Open-ReviewApp -Path 'C:\Review\MyApp.accdb'
Get-ReviewErrorLog -Mark
Open-ReviewForm frmMain
Save-ReviewCapture frmMain -Label before
Get-ReviewErrorLog
Close-ReviewApp
```

## Build the package

From the repository root, with Node.js installed:

```
npx @anthropic-ai/mcpb pack tools/review-mcp access-review.mcpb
```

That validates `manifest.json` and zips the folder (minus what `.mcpbignore` lists). When you change the tools, bump `version` in `manifest.json` and `$ServerVersion` in `AccessReviewMcp.ps1` together.

## In this repository's projects

Reviews of access-modern-ui PRs check all three palettes only when a PR adds controls or dialogs or changes theming or colours; layout, logic and data fixes are checked in Navy only. Any new error log entry is a must-fix finding unless explained.
