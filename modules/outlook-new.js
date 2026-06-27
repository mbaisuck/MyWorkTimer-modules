// outlook-new.js  —  New Outlook scanner module (contract: outlook-scanner-v1)
//
// PUBLISHED, UPDATABLE MODULE. Served from the public MyWorkTimer-modules repo,
// downloaded + SHA-256-verified + health-checked + cached by the module updater.
//
// This module is SELF-CONTAINED on purpose: it is loaded inside a locked vm
// sandbox that forbids require()/fs/child_process. So it carries everything it
// needs as plain data + pure functions:
//
//   - psScript      : the PowerShell UI-Automation script that reads New Outlook's
//                     open email (sender / recipients). LIFTED VERBATIM from the
//                     original scanner-new-outlook.js so behavior is identical.
//                     The BRAIN (not this module) spawns PowerShell with it; the
//                     sandbox itself never launches anything.
//   - parseLine(s)  : pure. One raw stdout line from that PowerShell -> a tagged
//                     object { kind:'diag'|'data'|'ignore', ... }.
//   - toScanResult(d): pure. A parsed 'data' object -> a contract ScanResult.
//   - scan(ctx)     : contract conformance for the health check. This scanner is
//                     push-driven (PowerShell streams lines), so a bare pull with
//                     no live line returns { status:'no-billable-item' }.
//
// When Microsoft changes New Outlook and breaks the scrape, THIS file is what
// gets re-published to the cloud. No app reinstall.

'use strict';

// ── Pure email helpers (normalization rules mirror outlook-scanner-v1) ────────
function normalizeEmail(s) {
    if (typeof s !== 'string') return null;
    const trimmed = s.trim().toLowerCase();
    if (trimmed.length === 0) return null;
    const at = trimmed.indexOf('@');
    if (at <= 0 || at !== trimmed.lastIndexOf('@') || at === trimmed.length - 1) {
        return null;
    }
    return trimmed;
}

function normalizeEmailArray(arr) {
    if (!Array.isArray(arr)) return [];
    const out = [];
    const seen = new Set();
    for (const raw of arr) {
        const e = normalizeEmail(raw);
        if (e !== null && !seen.has(e)) { seen.add(e); out.push(e); }
    }
    return out;
}

function safeString(s) { return typeof s === 'string' ? s : ''; }

// ── Line protocol parser ──────────────────────────────────────────────────────
// Mirrors how the original scanner-new-outlook.js read its PowerShell stdout:
//   "DEBUG|||<text>"                                  -> diagnostic heartbeat
//   "<title>|||<exe>|||<primary>|||<csv emails>|||<fromName>"  -> data
// Anything shorter is ignored.
function parseLine(line) {
    const parts = String(line == null ? '' : line).trim().split('|||');
    if (parts[0] === 'DEBUG') {
        return { kind: 'diag', text: parts.slice(1).join('|||') };
    }
    if (parts.length < 3) {
        return { kind: 'ignore' };
    }
    return {
        kind: 'data',
        title:   parts[0],
        exePath: parts[1],
        email:   parts[2] !== 'Unknown' ? parts[2] : null,
        allEmails: parts[3] ? parts[3].split(',') : [],
        fromName: parts[4] || '',
    };
}

// ── Interpreter: parsed 'data' -> contract ScanResult ────────────────────────
// Same logic as the Step-2 adapter, with the contract's invariants enforced by
// construction (normalized + deduped emails; meaningfulness rule).
function toScanResult(parsed) {
    if (!parsed || parsed.kind !== 'data') {
        return { status: 'no-billable-item' };
    }

    const allEmails = normalizeEmailArray(parsed.allEmails);
    let primary = null;
    if (typeof parsed.email === 'string' && parsed.email.length > 0) {
        primary = normalizeEmail(parsed.email);
    }
    const fromName = safeString(parsed.fromName);
    const title    = safeString(parsed.title);
    const exePath  = safeString(parsed.exePath);

    if (primary === null && fromName === '' && allEmails.length === 0) {
        return { status: 'no-billable-item' };
    }

    return {
        status: 'billable',
        item: {
            source: 'outlook-new',
            title,
            exePath,
            primary,
            allEmails,
            fromName,
            emailPayload: null, // UIA scrape cannot structure from/to/cc/bcc
        },
    };
}

// ── Contract conformance ──────────────────────────────────────────────────────
// Push-driven scanner: a synchronous pull with no buffered line is a no-op.
function scan(_ctx) {
    return { status: 'no-billable-item' };
}

// ── Embedded self-check (runs at load, INSIDE the sandbox) ────────────────────
// If the fragile parsing logic is broken, this throws -> the health check fails
// -> the updater quarantines the module and keeps the previous good one. This is
// how a bad cloud-pushed fix fails safe instead of silently mis-billing.
(function selfCheck() {
    const eq = (a, b) => JSON.stringify(a) === JSON.stringify(b);

    const d = parseLine('New Outlook|||olk|||jane@acme.com|||jane@acme.com,BOB@Acme.com , jane@acme.com|||Jane Smith');
    if (d.kind !== 'data') throw new Error('selfCheck: data line not parsed as data');

    const r = toScanResult(d);
    if (r.status !== 'billable') throw new Error('selfCheck: expected billable');
    if (r.item.source !== 'outlook-new') throw new Error('selfCheck: wrong source');
    if (r.item.primary !== 'jane@acme.com') throw new Error('selfCheck: wrong primary');
    if (!eq(r.item.allEmails, ['jane@acme.com', 'bob@acme.com'])) throw new Error('selfCheck: emails not normalized/deduped');
    if (r.item.emailPayload !== null) throw new Error('selfCheck: emailPayload must be null');

    const diag = parseLine('DEBUG|||foreground=explorer (not olk) idle');
    if (diag.kind !== 'diag') throw new Error('selfCheck: diag line not parsed as diag');

    const empty = toScanResult(parseLine('New Outlook|||olk|||Unknown|||| '));
    if (empty.status !== 'no-billable-item') throw new Error('selfCheck: empty scrape should be no-billable-item');

    if (scan({}).status !== 'no-billable-item') throw new Error('selfCheck: scan() must be a no-op pull');
})();

// ── The PowerShell screen-reading script (lifted verbatim) ────────────────────
const psScript = `
Add-Type -AssemblyName UIAutomationClient
Add-Type -AssemblyName UIAutomationTypes
Add-Type -TypeDefinition '
using System;
using System.Runtime.InteropServices;
public class Win32New {
    [DllImport("user32.dll")] public static extern IntPtr GetForegroundWindow();
    [DllImport("user32.dll")] public static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint lpdwProcessId);
}'

$emailRegex = [regex]'(?i)[a-z0-9._%+-]+@[a-z0-9.-]+\\.[a-z]{2,}'
$labelRegex = [regex]'(?i)^\\s*(from|to|cc|bcc|sender|sent|recipients?)\\b'

$AE  = [System.Windows.Automation.AutomationElement]
$CTP = $AE::ControlTypeProperty
$CTT = [System.Windows.Automation.ControlType]
$TSD = [System.Windows.Automation.TreeScope]::Descendants

# Cache Name + HelpText + ControlType + AutomationId in ONE pass so we are not making a
# cross-process UIA call per property per element (that would crawl in a WebView2 tree).
$cache = New-Object System.Windows.Automation.CacheRequest
$cache.Add($AE::NameProperty)
$cache.Add($AE::HelpTextProperty)
$cache.Add($AE::ControlTypeProperty)
$cache.Add($AE::AutomationIdProperty)

while ($true) {
    try {
        # ── Only act when New Outlook (olk) is the FOREGROUND window ──────────
        $hwnd = [Win32New]::GetForegroundWindow()
        $procId = 0
        [Win32New]::GetWindowThreadProcessId($hwnd, [ref]$procId) | Out-Null
        $fgName = "Unknown"
        if ($procId -gt 0) {
            $fp = Get-Process -Id $procId -ErrorAction SilentlyContinue
            if ($fp) { $fgName = $fp.ProcessName }
        }

        if ($fgName -ieq "olk") {
            # Bind UIA to the ACTUAL foreground window ($hwnd from GetForegroundWindow above), which
            # is by definition the olk window the user is looking at right now. The old code re-fetched
            # $olk.MainWindowHandle via Get-Process, which can go STALE (olk runs multiple processes /
            # the "main window" changes when a compose re-renders) — binding a dead handle, finding
            # nothing, and never recovering until Outlook restarts. Using $hwnd is always current and
            # self-heals when focus returns to a window.
            if ($hwnd -ne [IntPtr]::Zero) {
                $root = $AE::FromHandle($hwnd)

                # Look at the control types that actually carry addresses or labels.
                $cText = New-Object System.Windows.Automation.PropertyCondition($CTP, $CTT::Text)
                $cLink = New-Object System.Windows.Automation.PropertyCondition($CTP, $CTT::Hyperlink)
                $cEdit = New-Object System.Windows.Automation.PropertyCondition($CTP, $CTT::Edit)
                $cBtn  = New-Object System.Windows.Automation.PropertyCondition($CTP, $CTT::Button)
                $cond  = New-Object System.Windows.Automation.OrCondition(@($cText, $cLink, $cEdit, $cBtn))

                $els = $null
                $act = $cache.Activate()
                try { $els = $root.FindAll($TSD, $cond) } finally { $act.Dispose() }

                $found = New-Object System.Collections.Generic.List[string]
                $diag  = New-Object System.Collections.Generic.List[string]
                $diagCap = 16

                $fromName = ""
                foreach ($el in $els) {
                    $nm = ""; $ht = ""; $aid = ""
                    try { $nm  = [string]$el.Cached.Name }         catch { $nm  = "" }
                    try { $ht  = [string]$el.Cached.HelpText }      catch { $ht  = "" }
                    try { $aid = [string]$el.Cached.AutomationId }  catch { $aid = "" }
                    $hay = $nm + " " + $ht

                    # Capture the sender DISPLAY NAME from the reading-pane "From:" element.
                    # New Outlook hides a saved contact's EMAIL (shows only the display name), so
                    # this is frequently the only way to identify the sender. Prefer the element
                    # whose AutomationId ends in _FROM (the authoritative header field).
                    if ($nm -match '^\\s*From:\\s*(.+)$') {
                        $cand = ($matches[1] -replace '\\s*<.*$','' -replace '[<>]','').Trim()
                        if ($cand -ne '') {
                            if ($aid -match '_FROM$') { $fromName = $cand }
                            elseif ($fromName -eq '') { $fromName = $cand }
                        }
                    }

                    $ms = $emailRegex.Matches($hay)
                    if ($ms.Count -gt 0) {
                        foreach ($m in $ms) { $found.Add($m.Value.ToLower()) }
                        if ($diag.Count -lt $diagCap) {
                            $ct = "?"
                            try { $ct = [string]$el.Cached.ControlType.ProgrammaticName } catch {}
                            $diag.Add("EMAIL ct=$ct aid='$aid' name='$nm' help='$ht'")
                        }
                    } elseif ($labelRegex.IsMatch($nm)) {
                        if ($diag.Count -lt $diagCap) {
                            $diag.Add("LABEL aid='$aid' name='$nm' help='$ht'")
                        }
                    }
                }

                $uniq = @($found | Select-Object -Unique)
                $allStr  = ($uniq -join ",")
                $primary = if ($uniq.Count -gt 0) { $uniq[0] } else { "Unknown" }

                # Re-send EVERY poll while olk is in front (not only on change) so a transient
                # clear in the renderer can't permanently stop tracking. Emit if we have EITHER an
                # address OR a sender name (name-only = saved-contact sender New Outlook won't expose).
                if ($allStr -ne "" -or $fromName -ne "") {
                    Write-Output "New Outlook|||olk|||$primary|||$allStr|||$fromName"
                }
                Write-Output "DEBUG|||foreground=olk | scanned $($els.Count) els | emails=$($uniq.Count): $allStr | from='$fromName'"
                foreach ($d in $diag) { Write-Output "DEBUG|||$d" }

                # If we only got the account address (or nothing), dump the reading-pane text so we
                # can see WHERE the sender/recipient addresses are (which control + Name/HelpText).
                if ($uniq.Count -lt 2) {
                    $dump = 0
                    foreach ($el in $els) {
                        if ($dump -ge 25) { break }
                        $nm2 = ""; $ht2 = ""
                        try { $nm2 = [string]$el.Cached.Name }     catch {}
                        try { $ht2 = [string]$el.Cached.HelpText } catch {}
                        if ([string]::IsNullOrWhiteSpace($nm2) -and [string]::IsNullOrWhiteSpace($ht2)) { continue }
                        $ct2 = "?"
                        try { $ct2 = [string]$el.Cached.ControlType.ProgrammaticName } catch {}
                        Write-Output "DEBUG|||TXT ct=$ct2 name='$nm2' help='$ht2'"
                        $dump++
                    }
                }
            } else {
                Write-Output "DEBUG|||olk is foreground but no usable window handle (minimised?)"
            }
        } else {
            Write-Output "DEBUG|||foreground=$fgName (not olk) — idle"
        }
    } catch {
        Write-Output "DEBUG|||error: $($_.Exception.Message)"
    }

    Start-Sleep -Seconds 2
}
`;

module.exports = {
    contract: 'outlook-scanner-v1',
    id: 'outlook-new',
    source: 'outlook-new',
    version: '1.0.0',
    psScript,
    parseLine,
    toScanResult,
    scan,
};
