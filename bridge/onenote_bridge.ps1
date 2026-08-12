# OneNote COM bridge for the standalone OneNote MCP server.
#
# JSON-lines RPC over stdio: one request per stdin line, one response per
# stdout line. Runs as a persistent child of the Node host so the COM
# connection is created once.
#
#   Request:  {"id": 1, "op": "hierarchy", "args": {"scope": "sections"}}
#   Response: {"id": 1, "ok": true, "result": {...}}
#          or {"id": 1, "ok": false, "error": "..."}
#
# PowerShell (.NET COM interop), not Python: with x64 Click-to-Run Office,
# OneNote's typelib is registered only under the Win32 key, which breaks
# pywin32's typed dispatch AND dynamic GetIDsOfNames from 64-bit Python
# ("Library not registered"). .NET's late binder does not hit that path.

$ErrorActionPreference = "Stop"

# UTF-8 on BOTH stdio directions. The Node parent writes UTF-8; reading via
# [Console]::In uses the OEM codepage (CP437 on US systems), which mangles
# every non-ASCII character into mojibake (em-dash -> "ΓÇö") before it ever
# reaches OneNote. Explicit stream wrappers are deterministic regardless of
# console state; BOM-less so the first response line parses as clean JSON.
$script:Utf8NoBom = New-Object System.Text.UTF8Encoding($false)
$script:StdIn = New-Object System.IO.StreamReader([Console]::OpenStandardInput(), $script:Utf8NoBom)
$script:StdOut = New-Object System.IO.StreamWriter([Console]::OpenStandardOutput(), $script:Utf8NoBom)
$script:StdOut.AutoFlush = $true

$script:OneNS = "http://schemas.microsoft.com/office/onenote/2013/onenote"

function New-App {
  # New-Object binds to a running OneNote (Office desktop) or launches one.
  # Called lazily so the bridge starts even if OneNote isn't up yet, and again
  # on reconnect after OneNote was closed/restarted underneath us.
  $script:App = New-Object -ComObject OneNote.Application
}

$script:App = $null

# COM HRESULTs that mean "the OneNote we bound to is gone" — the caller should
# recreate the object and retry rather than surface a dead-handle error.
$script:DisconnectPatterns =
  'RPC server is unavailable|The object invoked has disconnected|been severed|' +
  '0x800706BA|0x80010108|0x800706BE'

function Invoke-Op($fn, $opArgs) {
  if ($null -eq $script:App) { New-App }
  try {
    return & $fn $opArgs
  } catch {
    $msg = "$($_.Exception.Message)"
    if ($msg -match $script:DisconnectPatterns) {
      # OneNote was closed/restarted; rebind to a live instance and retry once.
      New-App
      return & $fn $opArgs
    }
    throw
  }
}

$ScopeMap = @{ notebooks = 2; sections = 3; pages = 4 }

function New-NsManager([xml]$doc) {
  $ns = New-Object System.Xml.XmlNamespaceManager($doc.NameTable)
  $ns.AddNamespace("one", $script:OneNS)
  return $ns
}

function Strip-Html([string]$fragment) {
  if ($null -eq $fragment) { return "" }
  $text = $fragment -replace '<br[^>]*>', "`n" -replace '<[^>]+>', ''
  return [System.Net.WebUtility]::HtmlDecode($text)
}

function Node-Summary($el) {
  $out = [ordered]@{ kind = $el.LocalName; id = $el.GetAttribute("ID") }
  $name = $el.GetAttribute("name")
  if (-not $name) { $name = $el.GetAttribute("nickname") }
  $out.name = $name
  if ($el.GetAttribute("isCurrentlyViewed") -eq "true") { $out.currentlyViewed = $true }
  $lastModified = $el.GetAttribute("lastModifiedTime")
  if ($lastModified) { $out.lastModifiedTime = $lastModified }
  $dateTime = $el.GetAttribute("dateTime")
  if ($dateTime) { $out.dateTime = $dateTime }
  return $out
}

function Hierarchy-Tree($el) {
  $node = Node-Summary $el
  $children = @()
  foreach ($child in $el.ChildNodes) {
    if ($child.NodeType -ne "Element") { continue }
    if (@("Notebook", "SectionGroup", "Section", "Page") -contains $child.LocalName) {
      $children += , (Hierarchy-Tree $child)
    }
  }
  if ($children.Count -gt 0) { $node.children = $children }
  return $node
}

function Op-Ping($opArgs) { return [ordered]@{ pong = $true } }

function Op-Hierarchy($opArgs) {
  $scopeName = "pages"
  if ($opArgs.scope) { $scopeName = "$($opArgs.scope)".ToLower() }
  $scope = $ScopeMap[$scopeName]
  if ($null -eq $scope) { $scope = 4 }
  $startId = ""
  if ($opArgs.start_id) { $startId = $opArgs.start_id }
  $xmlOut = ""
  $script:App.GetHierarchy($startId, $scope, [ref]$xmlOut)
  [xml]$doc = $xmlOut
  $items = @()
  foreach ($el in $doc.DocumentElement.ChildNodes) {
    if ($el.NodeType -eq "Element") { $items += , (Hierarchy-Tree $el) }
  }
  # Scope "notebooks" returns the Notebooks root whose children are the
  # notebooks; other scopes nest the same way, so children of the document
  # element are always the items the caller asked about.
  return [ordered]@{ items = $items }
}

function Op-Get_Page($opArgs) {
  if (-not $opArgs.page_id) { throw "page_id is required" }
  $xmlOut = ""
  $script:App.GetPageContent($opArgs.page_id, [ref]$xmlOut)
  [xml]$doc = $xmlOut
  # GetElementsByTagName(localName, namespaceURI) avoids namespaced XPath +
  # XmlNamespaceManager entirely — the manager gets mangled when passed
  # through a helper function under Windows PowerShell, breaking the 2-arg
  # SelectSingleNode overload. Tag-name lookup is immune to that.
  $title = ""
  $titleEls = $doc.GetElementsByTagName("Title", $script:OneNS)
  if ($titleEls.Count -gt 0) {
    $titleT = $titleEls[0].GetElementsByTagName("T", $script:OneNS)
    if ($titleT.Count -gt 0) { $title = Strip-Html $titleT[0].InnerText }
  }
  $lines = @()
  foreach ($outline in $doc.GetElementsByTagName("Outline", $script:OneNS)) {
    foreach ($t in $outline.GetElementsByTagName("T", $script:OneNS)) {
      $lines += (Strip-Html $t.InnerText)
    }
  }
  return [ordered]@{
    page_id = $opArgs.page_id
    title   = $title
    text    = ($lines -join "`n")
  }
}

function Append-TextToPage([string]$pageId, [string]$text) {
  $xmlOut = ""
  $script:App.GetPageContent($pageId, [ref]$xmlOut)
  [xml]$doc = $xmlOut
  $outline = $doc.CreateElement("one", "Outline", $script:OneNS)
  $children = $doc.CreateElement("one", "OEChildren", $script:OneNS)
  [void]$outline.AppendChild($children)
  $paragraphs = $text -split "`n"
  foreach ($para in $paragraphs) {
    $oe = $doc.CreateElement("one", "OE", $script:OneNS)
    $t = $doc.CreateElement("one", "T", $script:OneNS)
    # OneNote interprets T CDATA as HTML, so plain text must be HTML-encoded
    # or literal "<task>"-style angle brackets get parsed as tags and vanish.
    [void]$t.AppendChild($doc.CreateCDataSection([System.Net.WebUtility]::HtmlEncode($para)))
    [void]$oe.AppendChild($t)
    [void]$children.AppendChild($oe)
  }
  [void]$doc.DocumentElement.AppendChild($outline)
  $script:App.UpdatePageContent($doc.OuterXml)
  return $paragraphs.Count
}

function Op-Append_Page($opArgs) {
  if (-not $opArgs.page_id) { throw "page_id is required" }
  if ($null -eq $opArgs.text) { throw "text is required" }
  $count = Append-TextToPage $opArgs.page_id $opArgs.text
  return [ordered]@{ appended_paragraphs = $count }
}

function Op-Create_Page($opArgs) {
  if (-not $opArgs.section_id) { throw "section_id is required" }
  $pageId = ""
  $script:App.CreateNewPage($opArgs.section_id, [ref]$pageId, 0)
  if ($opArgs.title) {
    $xmlOut = ""
    $script:App.GetPageContent($pageId, [ref]$xmlOut)
    [xml]$doc = $xmlOut
    # Tag-name lookup rather than namespaced XPath (see Op-Get_Page note).
    $titleEls = $doc.GetElementsByTagName("Title", $script:OneNS)
    if ($titleEls.Count -gt 0) {
      $titleEl = $titleEls[0]
    } else {
      $titleEl = $doc.CreateElement("one", "Title", $script:OneNS)
      [void]$doc.DocumentElement.PrependChild($titleEl)
    }
    $tEls = $titleEl.GetElementsByTagName("T", $script:OneNS)
    if ($tEls.Count -gt 0) {
      $t = $tEls[0]
    } else {
      $oe = $doc.CreateElement("one", "OE", $script:OneNS)
      $t = $doc.CreateElement("one", "T", $script:OneNS)
      [void]$oe.AppendChild($t)
      [void]$titleEl.AppendChild($oe)
    }
    $t.RemoveAll()
    [void]$t.AppendChild($doc.CreateCDataSection([System.Net.WebUtility]::HtmlEncode($opArgs.title)))
    $script:App.UpdatePageContent($doc.OuterXml)
  }
  if ($opArgs.body) { [void](Append-TextToPage $pageId $opArgs.body) }
  return [ordered]@{ page_id = $pageId }
}

function Op-Search($opArgs) {
  if (-not $opArgs.query) { throw "query is required" }
  $startId = ""
  if ($opArgs.start_id) { $startId = $opArgs.start_id }
  $xmlOut = ""
  $script:App.FindPages($startId, $opArgs.query, [ref]$xmlOut)
  [xml]$doc = $xmlOut
  # Tag-name lookup rather than namespaced XPath (see Op-Get_Page note).
  $pages = @()
  foreach ($page in $doc.GetElementsByTagName("Page", $script:OneNS)) {
    $pages += , ([ordered]@{ id = $page.GetAttribute("ID"); name = $page.GetAttribute("name") })
  }
  return [ordered]@{ query = $opArgs.query; pages = $pages }
}

function Op-Navigate($opArgs) {
  if (-not $opArgs.object_id) { throw "object_id is required" }
  $script:App.NavigateTo($opArgs.object_id, "", $false)
  return [ordered]@{ navigated = $opArgs.object_id }
}

function Remove-Object([string]$objectId, [bool]$permanent) {
  # deletePermanently=$false moves the object to the notebook's recycle bin
  # (recoverable) rather than erasing it. Works for any hierarchy object —
  # page, section, section group, or notebook.
  try {
    $script:App.DeleteHierarchy($objectId, $permanent)
  } catch {
    $msg = "$($_.Exception.Message)"
    if ($msg -match $script:DisconnectPatterns) { throw }
    # Some OneNote builds expose only the single-argument overload.
    $script:App.DeleteHierarchy($objectId)
  }
}

function Get-BoolArg($v) { if ($v) { return [bool]$v } return $false }

function Op-Delete_Page($opArgs) {
  if (-not $opArgs.page_id) { throw "page_id is required" }
  $permanent = Get-BoolArg $opArgs.permanent
  Remove-Object $opArgs.page_id $permanent
  return [ordered]@{ deleted = $opArgs.page_id; permanent = $permanent }
}

function Op-Delete_Section($opArgs) {
  if (-not $opArgs.section_id) { throw "section_id is required" }
  $permanent = Get-BoolArg $opArgs.permanent
  Remove-Object $opArgs.section_id $permanent
  return [ordered]@{ deleted = $opArgs.section_id; permanent = $permanent }
}

function Op-Delete_Notebook($opArgs) {
  if (-not $opArgs.notebook_id) { throw "notebook_id is required" }
  $permanent = Get-BoolArg $opArgs.permanent
  Remove-Object $opArgs.notebook_id $permanent
  return [ordered]@{ deleted = $opArgs.notebook_id; permanent = $permanent }
}

# --- Hierarchy helpers (create / rename / move / reorder) ---

$script:CftNotebook = 1
$script:CftFolder = 2   # section group
$script:CftSection = 3
$script:SlDefaultNotebookFolder = 2

function Find-HierNode([string]$startId, [int]$scope, [string]$localName, [string]$targetId) {
  # Fetch a fresh hierarchy slice and return the element whose ID matches.
  $h = ""
  $script:App.GetHierarchy($startId, $scope, [ref]$h)
  [xml]$doc = $h
  foreach ($el in $doc.GetElementsByTagName($localName, $script:OneNS)) {
    if ($el.GetAttribute("ID") -eq $targetId) { return $el }
  }
  return $null
}

function Op-Create_Section($opArgs) {
  # notebook_id may be a notebook OR a section group — OpenHierarchy resolves
  # the section path relative to it.
  $parent = $opArgs.notebook_id
  if (-not $parent) { $parent = $opArgs.parent_id }
  if (-not $parent) { throw "notebook_id (or parent_id) is required" }
  if (-not $opArgs.section_name) { throw "section_name is required" }
  $leaf = "$($opArgs.section_name)"
  if ($leaf -notmatch '\.one$') { $leaf = "$leaf.one" }
  $id = ""
  $script:App.OpenHierarchy($leaf, $parent, [ref]$id, $script:CftSection)
  return [ordered]@{ section_id = $id; name = $opArgs.section_name }
}

function Op-Create_Section_Group($opArgs) {
  $parent = $opArgs.parent_notebook_id
  if (-not $parent) { $parent = $opArgs.parent_section_group_id }
  if (-not $parent) { $parent = $opArgs.parent_id }
  if (-not $parent) { throw "parent_notebook_id or parent_section_group_id is required" }
  if (-not $opArgs.name) { throw "name is required" }
  $id = ""
  $script:App.OpenHierarchy("$($opArgs.name)", $parent, [ref]$id, $script:CftFolder)
  return [ordered]@{ section_group_id = $id; name = $opArgs.name }
}

function Op-Create_Notebook($opArgs) {
  if (-not $opArgs.name) { throw "name is required" }
  $folder = "$($opArgs.path)"
  if (-not $folder) {
    $loc = ""
    $script:App.GetSpecialLocation($script:SlDefaultNotebookFolder, [ref]$loc)
    $folder = $loc
  }
  $full = Join-Path $folder $opArgs.name
  $id = ""
  $script:App.OpenHierarchy($full, "", [ref]$id, $script:CftNotebook)
  return [ordered]@{ notebook_id = $id; name = $opArgs.name; path = $full }
}

function Op-Rename_Section($opArgs) {
  if (-not $opArgs.section_id) { throw "section_id is required" }
  if (-not $opArgs.new_name) { throw "new_name is required" }
  # Scan notebooks scope so the section node is found regardless of nesting.
  $sec = Find-HierNode "" 3 "Section" $opArgs.section_id
  if (-not $sec) { throw "section not found: $($opArgs.section_id)" }
  $sec.SetAttribute("name", $opArgs.new_name)
  $script:App.UpdateHierarchy($sec.OuterXml)
  return [ordered]@{ section_id = $opArgs.section_id; name = $opArgs.new_name }
}

function Op-Rename_Page($opArgs) {
  if (-not $opArgs.page_id) { throw "page_id is required" }
  if ($null -eq $opArgs.new_title) { throw "new_title is required" }
  Set-PageTitle $opArgs.page_id $opArgs.new_title
  return [ordered]@{ page_id = $opArgs.page_id; title = $opArgs.new_title }
}

function Set-PageTitle([string]$pageId, [string]$title) {
  $xmlOut = ""
  $script:App.GetPageContent($pageId, [ref]$xmlOut)
  [xml]$doc = $xmlOut
  $titleEls = $doc.GetElementsByTagName("Title", $script:OneNS)
  if ($titleEls.Count -gt 0) {
    $titleEl = $titleEls[0]
  } else {
    $titleEl = $doc.CreateElement("one", "Title", $script:OneNS)
    [void]$doc.DocumentElement.PrependChild($titleEl)
  }
  $tEls = $titleEl.GetElementsByTagName("T", $script:OneNS)
  if ($tEls.Count -gt 0) {
    $t = $tEls[0]
  } else {
    $oe = $doc.CreateElement("one", "OE", $script:OneNS)
    $t = $doc.CreateElement("one", "T", $script:OneNS)
    [void]$oe.AppendChild($t)
    [void]$titleEl.AppendChild($oe)
  }
  $t.RemoveAll()
  [void]$t.AppendChild($doc.CreateCDataSection([System.Net.WebUtility]::HtmlEncode($title)))
  $script:App.UpdatePageContent($doc.OuterXml)
}

function Get-SectionPageIds([string]$sectionId) {
  $h = ""
  $script:App.GetHierarchy($sectionId, 4, [ref]$h)
  [xml]$doc = $h
  $ids = @{}
  foreach ($pg in $doc.GetElementsByTagName("Page", $script:OneNS)) { $ids[$pg.GetAttribute("ID")] = $true }
  return $ids
}

# Object IDs are model-supplied (and can originate from untrusted note
# content), so any ID interpolated into hand-built XML must be escaped or a
# crafted "ID" breaks out of the attribute and injects hierarchy XML.
function Get-XmlAttrEscaped([string]$value) {
  return [System.Security.SecurityElement]::Escape($value)
}

# Magic-byte check for the image formats OneNote's Image element accepts.
function Test-SupportedImageHeader([byte[]]$bytes) {
  if ($bytes.Length -lt 8) { return $false }
  if ($bytes[0] -eq 0x89 -and $bytes[1] -eq 0x50 -and $bytes[2] -eq 0x4E -and $bytes[3] -eq 0x47 -and
      $bytes[4] -eq 0x0D -and $bytes[5] -eq 0x0A -and $bytes[6] -eq 0x1A -and $bytes[7] -eq 0x0A) { return $true } # PNG
  if ($bytes[0] -eq 0xFF -and $bytes[1] -eq 0xD8 -and $bytes[2] -eq 0xFF) { return $true }                         # JPEG
  if ($bytes[0] -eq 0x47 -and $bytes[1] -eq 0x49 -and $bytes[2] -eq 0x46 -and $bytes[3] -eq 0x38) { return $true } # GIF87a/89a
  if ($bytes[0] -eq 0x42 -and $bytes[1] -eq 0x4D) { return $true }                                                 # BMP
  if ($bytes[0] -eq 0x49 -and $bytes[1] -eq 0x49 -and $bytes[2] -eq 0x2A -and $bytes[3] -eq 0x00) { return $true } # TIFF LE
  if ($bytes[0] -eq 0x4D -and $bytes[1] -eq 0x4D -and $bytes[2] -eq 0x00 -and $bytes[3] -eq 0x2A) { return $true } # TIFF BE
  return $false
}

function Op-Move_Page($opArgs) {
  if (-not $opArgs.page_id) { throw "page_id is required" }
  if (-not $opArgs.target_section_id) { throw "target_section_id is required" }
  # OneNote regenerates a moved page's entire object ID (both the section GUID
  # and the page's own GUID change). Snapshot the target section's page IDs
  # before and after so we can hand back the new ID as the set difference.
  $before = Get-SectionPageIds $opArgs.target_section_id
  $xml =
    "<one:Section xmlns:one=`"$($script:OneNS)`" ID=`"$(Get-XmlAttrEscaped $opArgs.target_section_id)`">" +
    "<one:Page ID=`"$(Get-XmlAttrEscaped $opArgs.page_id)`" /></one:Section>"
  $script:App.UpdateHierarchy($xml)
  $after = Get-SectionPageIds $opArgs.target_section_id
  $newId = $null
  foreach ($k in $after.Keys) { if (-not $before.ContainsKey($k)) { $newId = $k; break } }
  return [ordered]@{
    page_id          = if ($newId) { $newId } else { $opArgs.page_id }
    previous_page_id = $opArgs.page_id
    section_id       = $opArgs.target_section_id
    id_changed       = [bool]$newId
  }
}

function Op-Move_Section($opArgs) {
  if (-not $opArgs.section_id) { throw "section_id is required" }
  if (-not $opArgs.target_parent_id) { throw "target_parent_id is required" }
  # The target parent is a notebook or a section group; reparent by submitting
  # it with the section as a child. Try SectionGroup wrapper first, then
  # Notebook, so either parent kind works.
  $secXml = "<one:Section ID=`"$(Get-XmlAttrEscaped $opArgs.section_id)`" />"
  $attempts = @(
    "<one:SectionGroup xmlns:one=`"$($script:OneNS)`" ID=`"$(Get-XmlAttrEscaped $opArgs.target_parent_id)`">$secXml</one:SectionGroup>",
    "<one:Notebook xmlns:one=`"$($script:OneNS)`" ID=`"$(Get-XmlAttrEscaped $opArgs.target_parent_id)`">$secXml</one:Notebook>"
  )
  $lastErr = $null
  foreach ($xml in $attempts) {
    try { $script:App.UpdateHierarchy($xml); return [ordered]@{ section_id = $opArgs.section_id; parent_id = $opArgs.target_parent_id } }
    catch { $lastErr = "$($_.Exception.Message)"; if ($lastErr -match $script:DisconnectPatterns) { throw } }
  }
  throw "move_section failed: $lastErr"
}

function Reorder-Children([string]$parentId, [int]$scope, [string]$childLocal, [string]$moveId, [string]$beforeId, [string]$afterId) {
  # Fetch the parent, reorder its child elements of type $childLocal so $moveId
  # lands before/after the reference, then resubmit the parent.
  $h = ""
  $script:App.GetHierarchy($parentId, $scope, [ref]$h)
  [xml]$doc = $h
  # The parent is the element whose ID == parentId (or the doc root for "").
  $parent = $null
  if ($parentId) {
    foreach ($el in $doc.GetElementsByTagName("*")) {
      if ($el.GetAttribute("ID") -eq $parentId) { $parent = $el; break }
    }
  }
  if (-not $parent) { $parent = $doc.DocumentElement }
  $kids = @()
  foreach ($c in $parent.ChildNodes) {
    if ($c.NodeType -eq "Element" -and $c.LocalName -eq $childLocal) { $kids += $c }
  }
  $moveNode = $kids | Where-Object { $_.GetAttribute("ID") -eq $moveId } | Select-Object -First 1
  if (-not $moveNode) { throw "$childLocal not found in parent: $moveId" }
  $refId = if ($beforeId) { $beforeId } else { $afterId }
  $refNode = $kids | Where-Object { $_.GetAttribute("ID") -eq $refId } | Select-Object -First 1
  if (-not $refNode) { throw "reference $childLocal not found: $refId" }
  [void]$parent.RemoveChild($moveNode)
  if ($beforeId) { [void]$parent.InsertBefore($moveNode, $refNode) }
  else { [void]$parent.InsertAfter($moveNode, $refNode) }
  $script:App.UpdateHierarchy($parent.OuterXml)
}

function Op-Reorder_Pages($opArgs) {
  if (-not $opArgs.page_id) { throw "page_id is required" }
  if (-not $opArgs.before_page_id -and -not $opArgs.after_page_id) {
    throw "before_page_id or after_page_id is required"
  }
  if (-not $opArgs.section_id) { throw "section_id is required (the section holding the pages)" }
  Reorder-Children $opArgs.section_id 4 "Page" $opArgs.page_id $opArgs.before_page_id $opArgs.after_page_id
  return [ordered]@{ page_id = $opArgs.page_id; reordered = $true }
}

function Op-Reorder_Sections($opArgs) {
  if (-not $opArgs.section_id) { throw "section_id is required" }
  if (-not $opArgs.before_section_id -and -not $opArgs.after_section_id) {
    throw "before_section_id or after_section_id is required"
  }
  if (-not $opArgs.parent_id) { throw "parent_id is required (the notebook or section group holding the sections)" }
  Reorder-Children $opArgs.parent_id 3 "Section" $opArgs.section_id $opArgs.before_section_id $opArgs.after_section_id
  return [ordered]@{ section_id = $opArgs.section_id; reordered = $true }
}

function Op-Update_Page($opArgs) {
  if (-not $opArgs.page_id) { throw "page_id is required" }
  if ($null -eq $opArgs.content) { throw "content is required" }
  $mode = "replace"
  if ($opArgs.mode) { $mode = "$($opArgs.mode)".ToLower() }
  if ($mode -eq "append") {
    $count = Append-TextToPage $opArgs.page_id $opArgs.content
    return [ordered]@{ page_id = $opArgs.page_id; mode = "append"; paragraphs = $count }
  }
  # replace: remove existing Outline objects, then append fresh content.
  $xmlOut = ""
  $script:App.GetPageContent($opArgs.page_id, [ref]$xmlOut)
  [xml]$doc = $xmlOut
  $outlineIds = @()
  foreach ($o in $doc.GetElementsByTagName("Outline", $script:OneNS)) {
    $oid = $o.GetAttribute("objectID")
    if ($oid) { $outlineIds += $oid }
  }
  foreach ($oid in $outlineIds) {
    try { $script:App.DeletePageContent($opArgs.page_id, $oid) } catch { }
  }
  $count = Append-TextToPage $opArgs.page_id $opArgs.content
  return [ordered]@{ page_id = $opArgs.page_id; mode = "replace"; paragraphs = $count }
}

# --- HTML -> OneNote XML translation -----------------------------------------
#
# OneNote's UpdatePageContent only accepts *inline* HTML (span/a) inside a
# one:T CDATA. Block structure — headings, paragraphs, lists, tables — must be
# expressed as one:OE / one:List / one:Table elements, or the whole update is
# rejected with 0x80042009 (invalid XML). These helpers walk a well-formed
# XHTML fragment and emit the corresponding OneNote XML.

function Prepare-XhtmlFragment([string]$html) {
  # [xml] parsing knows only the five XML entities; map the common HTML ones
  # to numeric references so fragments with &nbsp; etc. still parse.
  $s = $html
  $entityMap = @{
    '&nbsp;' = '&#160;'; '&mdash;' = '&#8212;'; '&ndash;' = '&#8211;'
    '&bull;' = '&#8226;'; '&hellip;' = '&#8230;'; '&middot;' = '&#183;'
    '&lsquo;' = '&#8216;'; '&rsquo;' = '&#8217;'; '&ldquo;' = '&#8220;'
    '&rdquo;' = '&#8221;'; '&copy;' = '&#169;'; '&reg;' = '&#174;'
    '&trade;' = '&#8482;'; '&rarr;' = '&#8594;'; '&larr;' = '&#8592;'
  }
  foreach ($k in $entityMap.Keys) { $s = $s.Replace($k, $entityMap[$k]) }
  return $s
}

function Protect-Cdata([string]$s) {
  # "]]>" inside CDATA would terminate the section; OneNote decodes the HTML
  # entity form back to the literal characters when rendering.
  return $s -replace '\]\]>', ']]&gt;'
}

function Convert-InlineHtml($node) {
  # Flatten an element's children into the inline HTML subset OneNote accepts
  # in T CDATA: <span style=...> and <a href=...>. Everything else is either
  # mapped to a styled span (b/i/u/em/strong/code) or unwrapped to its text.
  $sb = New-Object System.Text.StringBuilder
  foreach ($child in $node.ChildNodes) {
    if ($child.NodeType -eq "Text" -or $child.NodeType -eq "CDATA" -or $child.NodeType -eq "SignificantWhitespace") {
      [void]$sb.Append([System.Net.WebUtility]::HtmlEncode($child.Value))
      continue
    }
    if ($child.NodeType -ne "Element") { continue }
    $tag = $child.LocalName.ToLower()
    if ($tag -eq "ul" -or $tag -eq "ol") { continue } # nested lists handled at block level
    $inner = Convert-InlineHtml $child
    switch ($tag) {
      "b"      { [void]$sb.Append("<span style='font-weight:bold'>$inner</span>") }
      "strong" { [void]$sb.Append("<span style='font-weight:bold'>$inner</span>") }
      "i"      { [void]$sb.Append("<span style='font-style:italic'>$inner</span>") }
      "em"     { [void]$sb.Append("<span style='font-style:italic'>$inner</span>") }
      "u"      { [void]$sb.Append("<span style='text-decoration:underline'>$inner</span>") }
      "code"   { [void]$sb.Append("<span style='font-family:Consolas'>$inner</span>") }
      "br"     { [void]$sb.Append(" ") }
      "a"      {
        $href = [System.Net.WebUtility]::HtmlEncode($child.GetAttribute("href"))
        if ($href) { [void]$sb.Append("<a href=`"$href`">$inner</a>") }
        else { [void]$sb.Append($inner) }
      }
      "span"   {
        $style = $child.GetAttribute("style") -replace '"', "'"
        if ($style) { [void]$sb.Append("<span style='$style'>$inner</span>") }
        else { [void]$sb.Append($inner) }
      }
      default  { [void]$sb.Append($inner) }
    }
  }
  return $sb.ToString()
}

function Add-TextOE($doc, $parent, [string]$inlineHtml) {
  $oe = $doc.CreateElement("one", "OE", $script:OneNS)
  $t = $doc.CreateElement("one", "T", $script:OneNS)
  [void]$t.AppendChild($doc.CreateCDataSection((Protect-Cdata $inlineHtml)))
  [void]$oe.AppendChild($t)
  [void]$parent.AppendChild($oe)
  return $oe
}

function Add-ListItems($doc, $parent, $listNode, [bool]$ordered) {
  foreach ($li in $listNode.ChildNodes) {
    if ($li.NodeType -ne "Element" -or $li.LocalName.ToLower() -ne "li") { continue }
    $oe = $doc.CreateElement("one", "OE", $script:OneNS)
    $list = $doc.CreateElement("one", "List", $script:OneNS)
    if ($ordered) {
      $marker = $doc.CreateElement("one", "Number", $script:OneNS)
      $marker.SetAttribute("numberSequence", "0")
      $marker.SetAttribute("numberFormat", "##.")
      $marker.SetAttribute("language", "1033")
    } else {
      $marker = $doc.CreateElement("one", "Bullet", $script:OneNS)
      $marker.SetAttribute("bullet", "2")
      $marker.SetAttribute("fontSize", "11.0")
    }
    [void]$list.AppendChild($marker)
    [void]$oe.AppendChild($list)
    $t = $doc.CreateElement("one", "T", $script:OneNS)
    [void]$t.AppendChild($doc.CreateCDataSection((Protect-Cdata (Convert-InlineHtml $li))))
    [void]$oe.AppendChild($t)
    # A nested <ul>/<ol> inside the <li> becomes an OEChildren sub-list.
    $nested = $null
    foreach ($sub in $li.ChildNodes) {
      if ($sub.NodeType -eq "Element" -and ($sub.LocalName.ToLower() -eq "ul" -or $sub.LocalName.ToLower() -eq "ol")) {
        if ($null -eq $nested) {
          $nested = $doc.CreateElement("one", "OEChildren", $script:OneNS)
          [void]$oe.AppendChild($nested)
        }
        Add-ListItems $doc $nested $sub ($sub.LocalName.ToLower() -eq "ol")
      }
    }
    [void]$parent.AppendChild($oe)
  }
}

function Add-TableBlock($doc, $parent, $tableNode) {
  $rows = @()
  foreach ($el in $tableNode.GetElementsByTagName("tr")) { $rows += $el }
  if ($rows.Count -eq 0) { return }
  $maxCells = 0
  foreach ($row in $rows) {
    $n = 0
    foreach ($c in $row.ChildNodes) {
      if ($c.NodeType -eq "Element" -and ($c.LocalName.ToLower() -eq "td" -or $c.LocalName.ToLower() -eq "th")) { $n++ }
    }
    if ($n -gt $maxCells) { $maxCells = $n }
  }
  $oe = $doc.CreateElement("one", "OE", $script:OneNS)
  $table = $doc.CreateElement("one", "Table", $script:OneNS)
  $table.SetAttribute("bordersVisible", "true")
  $cols = $doc.CreateElement("one", "Columns", $script:OneNS)
  for ($i = 0; $i -lt $maxCells; $i++) {
    $col = $doc.CreateElement("one", "Column", $script:OneNS)
    $col.SetAttribute("index", "$i")
    $col.SetAttribute("width", "160")
    [void]$cols.AppendChild($col)
  }
  [void]$table.AppendChild($cols)
  foreach ($row in $rows) {
    $oneRow = $doc.CreateElement("one", "Row", $script:OneNS)
    $cellCount = 0
    foreach ($cell in $row.ChildNodes) {
      if ($cell.NodeType -ne "Element") { continue }
      $cellTag = $cell.LocalName.ToLower()
      if ($cellTag -ne "td" -and $cellTag -ne "th") { continue }
      $oneCell = $doc.CreateElement("one", "Cell", $script:OneNS)
      $cellKids = $doc.CreateElement("one", "OEChildren", $script:OneNS)
      $inline = Convert-InlineHtml $cell
      if ($cellTag -eq "th") { $inline = "<span style='font-weight:bold'>$inline</span>" }
      [void](Add-TextOE $doc $cellKids $inline)
      [void]$oneCell.AppendChild($cellKids)
      [void]$oneRow.AppendChild($oneCell)
      $cellCount++
    }
    # Pad short rows so every row has the declared column count.
    while ($cellCount -lt $maxCells) {
      $oneCell = $doc.CreateElement("one", "Cell", $script:OneNS)
      $cellKids = $doc.CreateElement("one", "OEChildren", $script:OneNS)
      [void](Add-TextOE $doc $cellKids "")
      [void]$oneCell.AppendChild($cellKids)
      [void]$oneRow.AppendChild($oneCell)
      $cellCount++
    }
    [void]$table.AppendChild($oneRow)
  }
  [void]$oe.AppendChild($table)
  [void]$parent.AppendChild($oe)
}

$script:HeadingSizes = @{ 1 = "20.0"; 2 = "16.0"; 3 = "13.0"; 4 = "12.0"; 5 = "11.0"; 6 = "11.0" }

function Add-HtmlBlock($doc, $parent, $node) {
  if ($node.NodeType -eq "Text" -or $node.NodeType -eq "CDATA") {
    if ("$($node.Value)".Trim()) {
      [void](Add-TextOE $doc $parent ([System.Net.WebUtility]::HtmlEncode($node.Value.Trim())))
    }
    return
  }
  if ($node.NodeType -ne "Element") { return }
  $tag = $node.LocalName.ToLower()
  if ($tag -match '^h([1-6])$') {
    $size = $script:HeadingSizes[[int]$Matches[1]]
    $inner = Convert-InlineHtml $node
    [void](Add-TextOE $doc $parent "<span style='font-size:${size}pt;font-weight:bold'>$inner</span>")
    return
  }
  switch ($tag) {
    "p"          { [void](Add-TextOE $doc $parent (Convert-InlineHtml $node)) }
    "blockquote" { [void](Add-TextOE $doc $parent "<span style='font-style:italic'>$(Convert-InlineHtml $node)</span>") }
    "ul"         { Add-ListItems $doc $parent $node $false }
    "ol"         { Add-ListItems $doc $parent $node $true }
    "table"      { Add-TableBlock $doc $parent $node }
    "pre"        {
      foreach ($codeLine in ("$($node.InnerText)" -split "`r?`n")) {
        $encoded = [System.Net.WebUtility]::HtmlEncode($codeLine)
        [void](Add-TextOE $doc $parent "<span style='font-family:Consolas'>$encoded</span>")
      }
    }
    "hr"         { [void](Add-TextOE $doc $parent "") }
    default      {
      # div/section/article and unknown containers: recurse block children.
      foreach ($child in $node.ChildNodes) { Add-HtmlBlock $doc $parent $child }
    }
  }
}

function Op-Insert_Rich_Content($opArgs) {
  if (-not $opArgs.page_id) { throw "page_id is required" }
  $xmlOut = ""
  $script:App.GetPageContent($opArgs.page_id, [ref]$xmlOut)
  [xml]$doc = $xmlOut
  $outline = $doc.CreateElement("one", "Outline", $script:OneNS)
  $children = $doc.CreateElement("one", "OEChildren", $script:OneNS)
  [void]$outline.AppendChild($children)
  if ($opArgs.html) {
    $prepared = Prepare-XhtmlFragment "$($opArgs.html)"
    try {
      [xml]$frag = "<root>$prepared</root>"
    } catch {
      throw ("html must be a well-formed XHTML fragment (self-close void tags, " +
        "match every open tag): $($_.Exception.Message)")
    }
    foreach ($child in $frag.DocumentElement.ChildNodes) {
      Add-HtmlBlock $doc $children $child
    }
    if (-not $children.HasChildNodes -and -not $opArgs.image_path) {
      throw "html fragment produced no content (empty or unsupported elements only)"
    }
  }
  if ($opArgs.image_path) {
    if (-not (Test-Path -LiteralPath $opArgs.image_path)) { throw "image_path not found: $($opArgs.image_path)" }
    # Embedded images end up in notebooks that sync to OneDrive and may be
    # shared, so image_path must actually be an image — otherwise this op is an
    # arbitrary-file-read that copies local files off the machine.
    $imgInfo = Get-Item -LiteralPath $opArgs.image_path
    if ($imgInfo.Length -gt 26214400) { throw "image_path exceeds the 25 MB embed limit: $($opArgs.image_path)" }
    $bytes = [System.IO.File]::ReadAllBytes($opArgs.image_path)
    if (-not (Test-SupportedImageHeader $bytes)) {
      throw "image_path is not a supported image (png, jpeg, gif, bmp, or tiff): $($opArgs.image_path)"
    }
    $b64 = [System.Convert]::ToBase64String($bytes)
    $oe = $doc.CreateElement("one", "OE", $script:OneNS)
    $img = $doc.CreateElement("one", "Image", $script:OneNS)
    $img.SetAttribute("format", "auto")
    $data = $doc.CreateElement("one", "Data", $script:OneNS)
    [void]$data.AppendChild($doc.CreateTextNode($b64))
    [void]$img.AppendChild($data)
    [void]$oe.AppendChild($img)
    [void]$children.AppendChild($oe)
  }
  if (-not $children.HasChildNodes) { throw "provide html and/or image_path" }
  [void]$doc.DocumentElement.AppendChild($outline)
  $script:App.UpdatePageContent($doc.OuterXml)
  return [ordered]@{ page_id = $opArgs.page_id; inserted = $true }
}

$script:PubFormat = @{ onenote = 0; package = 1; mhtml = 2; pdf = 3; xps = 4; word = 5; docx = 5; emf = 6; html = 7 }

function Op-Export($opArgs) {
  if (-not $opArgs.object_id) { throw "object_id is required" }
  if (-not $opArgs.target_path) { throw "target_path is required" }
  $fmtName = "pdf"
  if ($opArgs.format) { $fmtName = "$($opArgs.format)".ToLower() }
  $fmt = $script:PubFormat[$fmtName]
  if ($null -eq $fmt) { throw "unknown export format: $fmtName (pdf|html|docx|mhtml|xps|onenote)" }
  $script:App.Publish($opArgs.object_id, $opArgs.target_path, $fmt, "")
  return [ordered]@{ object_id = $opArgs.object_id; path = $opArgs.target_path; format = $fmtName }
}

while ($null -ne ($line = $script:StdIn.ReadLine())) {
  $line = $line.Trim()
  if (-not $line) { continue }
  $reqId = $null
  try {
    $req = $line | ConvertFrom-Json
    $reqId = $req.id
    $fn = Get-Command -Name ("Op-" + $req.op) -CommandType Function -ErrorAction SilentlyContinue
    if (-not $fn) { throw "unknown op: $($req.op)" }
    $result = Invoke-Op $fn $req.args
    $resp = [ordered]@{ id = $reqId; ok = $true; result = $result }
  } catch {
    $resp = [ordered]@{ id = $reqId; ok = $false; error = "$($_.Exception.Message)" }
  }
  $script:StdOut.WriteLine(($resp | ConvertTo-Json -Depth 20 -Compress))
}
