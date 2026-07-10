# Syncfusion Evaluation

Two different Syncfusion-based approaches were evaluated.

Although both use **Syncfusion XlsIO** on the backend to read Excel workbooks, they differ significantly in architecture:

- **Option F** uses Syncfusion's complete spreadsheet platform, including the React Spreadsheet component.
- **Option G** uses Syncfusion purely as a high-fidelity Excel parser and exports a custom JSON model rendered by the application's own React viewer.

## F — Syncfusion Spreadsheet

### Architecture

```text
Excel Workbook
       │
       ▼
Syncfusion XlsIO (.NET)
       │
       ▼
Workbook.Open(...)
       │
       ▼
Syncfusion Spreadsheet JSON
       │
       ▼
Syncfusion React Spreadsheet
```

### Backend processing

The backend:

1. Opens the workbook with XlsIO.
2. Enables formula calculation.
3. Replaces formulas with calculated values.
4. Preserves table filtering metadata.
5. Converts the workbook into the JSON format expected by the Syncfusion Spreadsheet component.

Example:

```csharp
var workbook = appExcel.Workbooks.Open(stream);

foreach (var sheet in workbook.Worksheets)
{
    sheet.EnableSheetCalculations();

    foreach (var cell in sheet.UsedRange.Cells)
    {
        if (!string.IsNullOrEmpty(cell.Formula))
        {
            var value = cell.CalculatedValue;

            cell.Clear(ExcelClearOptions.ClearContent);
            cell.Value = value;
        }
    }
}

var result = Workbook.Open(openRequest);
```

### Advantages

- Highest Excel fidelity of all evaluated options.
- Native support for formulas, tables, filters, sorting, conditional formatting, merged cells, images, comments, freeze panes and workbook navigation.
- Very little custom frontend code.
- Vendor maintains spreadsheet behavior.
- New Excel features often require configuration rather than custom development.

### Disadvantages

- Commercial licensing.
- Larger browser footprint than all custom-viewer approaches.
- Strong vendor lock-in.
- Spreadsheet component is fundamentally designed for editing and must be configured as read-only.
- Less control over user experience and visual integration.

### Best fit

Recommended when the goal is:

> Provide a browser experience that behaves as closely as possible to Microsoft Excel.

---

# G — Syncfusion XlsIO + Custom React Viewer

### Architecture

```text
Excel Workbook
       │
       ▼
Syncfusion XlsIO (.NET)
       │
       ▼
Custom DTO Model
       │
       ▼
JSON
       │
       ▼
Custom React Viewer
```

### Backend extraction

The workbook is parsed using Syncfusion XlsIO and converted into a custom JSON structure consumed by the frontend.

The exported model currently includes:

- Cell values
- Cell styles
- Fonts
- Borders
- Alignment
- Foreground/background colors
- Charts
- Images
- Frozen panes
- Merged cells

### Example DTO

```text
WorksheetDto
├─ Grid
├─ Charts
├─ Images
├─ FrozenRows
├─ FrozenColumns
└─ MergedCells
```

## Chart support

A notable advantage of this approach is chart extraction.

Charts are exposed through the Syncfusion chart API and rendered as PNG images.

```csharp
foreach (IChart chart in worksheet.Charts)
{
    chart.SaveAsImage(chartStream);
}
```

Unlike the ClosedXML, ExcelJS, and OSS Univer approaches, workbook charts are therefore available to the viewer.

For a read-only preview scenario this is often sufficient because users see the exact chart as rendered by Excel.

### Advantages

- Lightweight frontend.
- Full control of the rendering experience.
- No spreadsheet engine executed in the browser.
- Native chart extraction.
- Native image extraction.
- Native merged-cell support.
- Native freeze-pane support.
- Existing React viewer can be reused.
- Lower vendor lock-in than the Syncfusion Spreadsheet approach.

### Disadvantages

- Commercial backend dependency.
- Frontend rendering must still be maintained.
- Every new Excel feature requires development.
- Conditional-formatting behavior remains custom.
- Table rendering remains custom.
- Charts are rendered as images rather than interactive objects.

### Best fit

Recommended when the goal is:

> Deliver a lightweight, high-fidelity read-only Excel preview integrated into an existing React application.

---

# Comparison

| Capability | Syncfusion Spreadsheet | Syncfusion XlsIO + Custom Viewer |
|------------|----------------------|----------------------------------|
| Excel fidelity | Excellent | Very good |
| Charts | Native | PNG rendering |
| Images | Native | Native |
| Merged cells | Native | Native |
| Freeze panes | Native | Native |
| Tables | Native | Custom rendering |
| Conditional formatting | Native | Custom rendering |
| Filtering | Native | Requires custom implementation |
| Keyboard navigation | Native | Custom implementation |
| Browser footprint | Larger | Smaller |
| Frontend control | Limited | Full |
| Vendor lock-in | High | Moderate |
| Development effort | Low | Medium/High |
| Long-term maintenance | Lower | Higher |

---

# Updated Recommendation

The addition of Syncfusion changes the conclusions of the original evaluation.

Previously, chart rendering was the primary remaining feature gap across all open-source approaches.

With Syncfusion XlsIO, charts can be extracted and rendered, removing the largest functional limitation identified in the earlier POC.

## If Excel fidelity is the primary objective

Choose:

```text
Syncfusion Spreadsheet
```

This provides the closest experience to desktop Excel while minimizing custom development.

## If a lightweight read-only preview is the primary objective

Choose:

```text
Syncfusion XlsIO
+
Custom JSON model
+
Custom React viewer
```

This approach preserves the lightweight architecture of the original backend-parsed solutions while adding native chart support and maintaining full control over the user experience.

---

# Final Ranking

| Rank | Solution |
|--------|----------|
| 🥇 | Syncfusion XlsIO → JSON → Custom Viewer |
| 🥈 | Syncfusion Spreadsheet |
| 🥉 | ClosedXML → JSON → Custom Grid |
| 4 | Jspreadsheet CE |
| 5 | react-data-grid |
| 6 | ExcelJS |
| 7 | Univer OSS |

## Rationale

The Syncfusion XlsIO + Custom Viewer architecture combines:

- high Excel fidelity,
- chart rendering,
- lightweight browser footprint,
- full UI ownership,
- and a read-only-first design.

It therefore provides the strongest balance between capability, maintainability, and user experience for the specific goal of an embedded Excel preview component.
