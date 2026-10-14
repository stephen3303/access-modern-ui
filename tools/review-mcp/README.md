# Access Review MCP server

Lets a Claude Code session do visual review of an Access front end without using the desktop. It starts its own Access instance over COM, parks its windows past the edge of the screen, and captures forms with `PrintWindow`, which asks Access to draw a window into a bitmap. Nothing comes to the front, nothing takes focus, and no mouse or keyboard input is sent, so someone can keep working on the same machine. In off-screen mode a watcher keeps every Access window, including popups the app opens or re-centres on its own, past the desktop edge for the whole session; a newly shown popup can appear for a few milliseconds before it is moved.

```
AccessReview.ps1      the review functions (usable from plain PowerShell too)
AccessReviewMcp.ps1   the MCP server (stdio) that exposes them as tools
```

## Tools

| Tool | What it does |
|---|---|
| `review_lock` / `review_unlock` | Takes and releases the shared reviewer lock files listed in `ACCESS_REVIEW_LOCKS`, re-reading after 3 s to catch a race. Releases only its own. |
| `open_app` | Opens an existing front end in a new Access instance (never creates a file). `mode`: `offscreen` (default) or `onscreen` for a dedicated review machine. |
| `error_log` | `mark=true` before testing, then `mark=false` to list new `error_log` rows and new `errors.log` text. |
| `compile` | Compiles and saves all modules. |
| `run` | Runs a public procedure with `Application.Run`, e.g. SABER `Login`. Whole numbers are passed as Long (Double when too large) and output arguments come back. |
| `set_palette` | Navy, Midnight or Teal Navy (`SetPalette` 1, 2, 6), then closes and reopens the open forms at their last size, as the app's Appearance switch does. A form left open keeps part of the old drawing (pills and progress bars went missing in Midnight). |
| `open_form` | Opens a form and sizes it to its design size (form width, visible section heights), or to `width_px` x `height_px` at 96 dpi. Reports whether the Access window clips it. |
| `resize_form` | Design size, a given size (small check about 1100x700) or maximized (Access window at the main display's work area, as a real maximized window gets). A shell form that fills the Access window, such as kit frmMain, is sized by resizing the Access window until its inside size matches; `SizeMatches` reports whether it did. |
| `capture` | PNG of a form, or a zoomed crop, taken once two frames in a row match (so a form still drawing isn't caught mid-draw; `Stable` reports it). Returned as an image and saved in the output folder, named form-palette-label. |
| `controls` | Each control's Left/Top/Width/Height (96 dpi px), section, caption and colours, for measuring gaps and the 16 px card inset. |
| `dialogs` | Lists and captures open message boxes and popups; `close=true` closes them. Also lists dialogs seen during earlier calls. |
| `close_app` | Resets the palette to Navy, quits without saving design changes and confirms the process is gone. |

Every Access call runs under a guard. A window that appears during a call is parked off-screen, captured and described. VBA and Office message boxes are closed so the call returns. A Yes/No box ignores Close, so the guard presses No, or the default button when there is no No. Another window (a kit dialog opened modally) is closed if it is still blocking the call after 4 s, except while `open_app` starts the front end, so a splash or sign-in form stays open. A call that runs past its timeout ends Access. The tool's result lists what the guard saw.

PowerShell can't read Access form, control and section properties directly (their type information fails, so `$form.Hwnd` comes back empty). The tools read them with `Application.Eval("Forms![name]...")` and set the form's inside size late-bound.

## Setup

Register it with Claude Code on the review machine, in the project's `.mcp.json` or with `claude mcp add`:

```json
{
  "mcpServers": {
    "access-review": {
      "command": "powershell.exe",
      "args": ["-NoProfile", "-ExecutionPolicy", "Bypass", "-File", "C:\\path\\to\\access-modern-ui\\tools\\review-mcp\\AccessReviewMcp.ps1"],
      "env": { "ACCESS_REVIEW_LOCKS": "C:\\...\\_review\\ACCESS_IN_USE.txt;C:\\...\\_review\\kit\\LOCK.txt" }
    }
  }
}
```

It uses Windows PowerShell 5.1 and .NET Framework, which come with Windows, plus Access. Nothing else is installed.

## Typical review run

1. `review_lock` with the thread name.
2. `open_app` on the approved scratch front end, `error_log` with `mark=true`.
3. For each changed page: `open_form` at its designed size, `capture`, `controls` to measure spacing, `dialogs`.
4. Palettes only when the PR adds controls or dialogs or changes colours: `set_palette`, `capture` again.
5. `resize_form` small and maximized, `capture`.
6. `error_log` (new entries are must-fix unless explained), `close_app`, `review_unlock`.

## Limits

- Hover effects, cursor changes and flicker during a real click can't be seen without a real pointer. Check those by hand.
- `PrintWindow` can't draw a minimized window, and a locked or disconnected Windows session draws nothing. The capture warns when an image comes back one solid colour. On a server, keep the review session unlocked: use auto-logon, or end Remote Desktop with `tscon <session id> /dest:console` instead of disconnecting.
- A window can't be bigger than the desktop, so the Access window is at most the desktop size (default 1920x1200). On a server with a small default resolution, raise it, or forms larger than the Access window will be clipped (`ClippedByAccessWindow` says so).
- Under automation, Access runs macros without the Enable Content bar. Still open only the approved scratch file, as the review rules say.
- Typing into fields and clicking buttons on a form isn't included yet. Use `run` for public procedures in the meantime.
