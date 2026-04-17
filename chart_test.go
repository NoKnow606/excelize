package excelize

import (
	"bytes"
	"encoding/xml"
	"fmt"
	"path/filepath"
	"testing"

	"github.com/stretchr/testify/assert"
)

func TestChartSize(t *testing.T) {
	f := NewFile()
	sheet1 := f.GetSheetName(0)

	categories := map[string]string{
		"A2": "Small",
		"A3": "Normal",
		"A4": "Large",
		"B1": "Apple",
		"C1": "Orange",
		"D1": "Pear",
	}
	for cell, v := range categories {
		assert.NoError(t, f.SetCellValue(sheet1, cell, v))
	}

	values := map[string]int{
		"B2": 2,
		"C2": 3,
		"D2": 3,
		"B3": 5,
		"C3": 2,
		"D3": 4,
		"B4": 6,
		"C4": 7,
		"D4": 8,
	}
	for cell, v := range values {
		assert.NoError(t, f.SetCellValue(sheet1, cell, v))
	}

	assert.NoError(t, f.AddChart("Sheet1", "E4", &Chart{
		Type: Col3DClustered,
		Dimension: ChartDimension{
			Width:  640,
			Height: 480,
		},
		Series: []ChartSeries{
			{Name: "Sheet1!$A$2", Categories: "Sheet1!$B$1:$D$1", Values: "Sheet1!$B$2:$D$2"},
			{Name: "Sheet1!$A$3", Categories: "Sheet1!$B$1:$D$1", Values: "Sheet1!$B$3:$D$3"},
			{Name: "Sheet1!$A$4", Categories: "Sheet1!$B$1:$D$1", Values: "Sheet1!$B$4:$D$4"},
		},
		Title: []RichTextRun{{Text: "3D Clustered Column Chart"}},
	}))

	var buffer bytes.Buffer

	// Save spreadsheet by the given path.
	assert.NoError(t, f.Write(&buffer))

	newFile, err := OpenReader(&buffer)
	assert.NoError(t, err)

	chartsNum := newFile.countCharts()
	if !assert.Equal(t, 1, chartsNum, "Expected 1 chart, actual %d", chartsNum) {
		t.FailNow()
	}

	var (
		workdir decodeWsDr
		anchor  decodeCellAnchor
	)

	content, ok := newFile.Pkg.Load("xl/drawings/drawing1.xml")
	assert.True(t, ok, "Can't open the chart")

	err = xml.Unmarshal(content.([]byte), &workdir)
	if !assert.NoError(t, err) {
		t.FailNow()
	}

	err = xml.Unmarshal([]byte("<decodeCellAnchor>"+
		workdir.TwoCellAnchor[0].Content+"</decodeCellAnchor>"), &anchor)
	if !assert.NoError(t, err) {
		t.FailNow()
	}

	if !assert.Equal(t, 4, anchor.From.Col, "Expected 'from' column 4") ||
		!assert.Equal(t, 3, anchor.From.Row, "Expected 'from' row 3") {

		t.FailNow()
	}

	assert.Equal(t, 14, anchor.To.Col, "Expected 'to' column 14")
	assert.Equal(t, 29, anchor.To.Row, "Expected 'to' row 29")
}

func TestAddDrawingChart(t *testing.T) {
	f := NewFile()
	assert.EqualError(t, f.addDrawingChart("SheetN", "", "", 0, 0, 0, nil), newCellNameToCoordinatesError("", newInvalidCellNameError("")).Error())

	path := "xl/drawings/drawing1.xml"
	f.Pkg.Store(path, MacintoshCyrillicCharset)
	assert.EqualError(t, f.addDrawingChart("Sheet1", path, "A1", 0, 0, 0, &GraphicOptions{PrintObject: boolPtr(true), Locked: boolPtr(false)}), "XML syntax error on line 1: invalid UTF-8")
}

func TestAddSheetDrawingChart(t *testing.T) {
	f := NewFile()
	path := "xl/drawings/drawing1.xml"
	f.Pkg.Store(path, MacintoshCyrillicCharset)
	assert.EqualError(t, f.addSheetDrawingChart(path, 0, &GraphicOptions{PrintObject: boolPtr(true), Locked: boolPtr(false)}), "XML syntax error on line 1: invalid UTF-8")
}

func TestDeleteDrawing(t *testing.T) {
	f := NewFile()
	path := "xl/drawings/drawing1.xml"
	f.Pkg.Store(path, MacintoshCyrillicCharset)
	_, err := f.deleteDrawing(0, 0, path, "Chart")
	assert.EqualError(t, err, "XML syntax error on line 1: invalid UTF-8")
	f, err = OpenFile(filepath.Join("test", "Book1.xlsx"))
	assert.NoError(t, err)
	f.Drawings.Store(path, &xlsxWsDr{OneCellAnchor: []*xdrCellAnchor{{
		GraphicFrame: string(MacintoshCyrillicCharset),
	}}})
	_, err = f.deleteDrawing(0, 0, path, "Chart")
	assert.EqualError(t, err, "XML syntax error on line 1: invalid UTF-8")
	f.Drawings.Store(path, &xlsxWsDr{TwoCellAnchor: []*xdrCellAnchor{{
		GraphicFrame: string(MacintoshCyrillicCharset),
	}}})
	_, err = f.deleteDrawing(0, 0, path, "Chart")
	assert.EqualError(t, err, "XML syntax error on line 1: invalid UTF-8")
}

func TestAddChart(t *testing.T) {
	f, err := OpenFile(filepath.Join("test", "Book1.xlsx"))
	if !assert.NoError(t, err) {
		t.FailNow()
	}

	categories := map[string]string{"A30": "SS", "A31": "S", "A32": "M", "A33": "L", "A34": "LL", "A35": "XL", "A36": "XXL", "A37": "XXXL", "B29": "Apple", "C29": "Orange", "D29": "Pear"}
	values := map[string]int{"B30": 1, "C30": 1, "D30": 1, "B31": 2, "C31": 2, "D31": 2, "B32": 3, "C32": 3, "D32": 3, "B33": 4, "C33": 4, "D33": 4, "B34": 5, "C34": 5, "D34": 5, "B35": 6, "C35": 6, "D35": 6, "B36": 7, "C36": 7, "D36": 7, "B37": 8, "C37": 8, "D37": 8}
	for k, v := range categories {
		assert.NoError(t, f.SetCellValue("Sheet1", k, v))
	}
	for k, v := range values {
		assert.NoError(t, f.SetCellValue("Sheet1", k, v))
	}
	assert.EqualError(t, f.AddChart("Sheet1", "P1", nil), ErrParameterInvalid.Error())

	// Test add chart on not exists worksheet
	assert.EqualError(t, f.AddChart("SheetN", "P1", nil), "sheet SheetN does not exist")
	maximum, minimum, zero := 7.5, 0.5, .0
	series := []ChartSeries{
		{Name: "Sheet1!$A$30", Categories: "Sheet1!$B$29:$D$29", Values: "Sheet1!$B$30:$D$30"},
		{Name: "Sheet1!$A$31", Categories: "Sheet1!$B$29:$D$29", Values: "Sheet1!$B$31:$D$31"},
		{Name: "Sheet1!$A$32", Categories: "Sheet1!$B$29:$D$29", Values: "Sheet1!$B$32:$D$32"},
		{Name: "Sheet1!$A$33", Categories: "Sheet1!$B$29:$D$29", Values: "Sheet1!$B$33:$D$33"},
		{Name: "Sheet1!$A$34", Categories: "Sheet1!$B$29:$D$29", Values: "Sheet1!$B$34:$D$34"},
		{Name: "Sheet1!$A$35", Categories: "Sheet1!$B$29:$D$29", Values: "Sheet1!$B$35:$D$35"},
		{Name: "Sheet1!$A$36", Categories: "Sheet1!$B$29:$D$29", Values: "Sheet1!$B$36:$D$36"},
		{
			Name: "Sheet1!$A$37", Categories: "Sheet1!$B$29:$D$29", Values: "Sheet1!$B$37:$D$37",
			Marker: ChartMarker{
				Fill: Fill{Type: "pattern", Color: []string{"FFFFFF"}, Pattern: 1},
				Border: ChartLine{
					Type:  ChartLineSolid,
					Dash:  ChartDashDot,
					Fill:  Fill{Type: "pattern", Color: []string{"#FFFF00"}, Pattern: 1},
					Width: 1,
				},
			},
			Legend: ChartLegend{Font: &Font{Family: "Arial", Size: 11, Strike: true, Color: "777777"}},
		},
	}
	series2 := []ChartSeries{
		{
			Name: "Sheet1!$A$30", Categories: "Sheet1!$B$29:$D$29", Values: "Sheet1!$B$30:$D$30",
			Fill:   Fill{Type: "pattern", Color: []string{"000000"}, Pattern: 1, Transparency: 60},
			Marker: ChartMarker{Symbol: "none", Size: 10},
		},
		{Name: "Sheet1!$A$31", Categories: "Sheet1!$B$29:$D$29", Values: "Sheet1!$B$31:$D$31"},
		{Name: "Sheet1!$A$32", Categories: "Sheet1!$B$29:$D$29", Values: "Sheet1!$B$32:$D$32"},
		{Name: "Sheet1!$A$33", Categories: "Sheet1!$B$29:$D$29", Values: "Sheet1!$B$33:$D$33"},
		{Name: "Sheet1!$A$34", Categories: "Sheet1!$B$29:$D$29", Values: "Sheet1!$B$34:$D$34"},
		{Name: "Sheet1!$A$35", Categories: "Sheet1!$B$29:$D$29", Values: "Sheet1!$B$35:$D$35"},
		{Name: "Sheet1!$A$36", Categories: "Sheet1!$B$29:$D$29", Values: "Sheet1!$B$36:$D$36"},
		{Name: "Sheet1!$A$37", Categories: "Sheet1!$B$29:$D$29", Values: "Sheet1!$B$37:$D$37", Line: ChartLine{Dash: ChartDashDash, Width: 0.25}},
	}
	series3 := []ChartSeries{{Name: "Sheet1!$A$30", Categories: "Sheet1!$A$30:$D$37", Values: "Sheet1!$B$30:$B$37"}}
	series4 := []ChartSeries{{Name: "Sheet1!$A$30", Categories: "Sheet1!$A$30:$D$37", Values: "Sheet1!$B$30:$B$37", DataPoint: []ChartDataPoint{
		{Index: 0, Fill: Fill{Type: "pattern", Color: []string{"003F5C"}, Pattern: 1}},
		{Index: 1, Fill: Fill{Type: "pattern", Color: []string{"58508D"}, Pattern: 1}},
		{Index: 2, Fill: Fill{Type: "pattern", Color: []string{"BC5090"}, Pattern: 1}},
		{Index: 3, Fill: Fill{Type: "pattern", Color: []string{"FF6361"}, Pattern: 1}},
	}}}
	series5 := []ChartSeries{
		{Name: "Sheet1!$A$30", Categories: "Sheet1!$B$29:$D$29", Values: "Sheet1!$B$30:$D$30", Sizes: "Sheet1!$B$30:$D$30", DataLabelPosition: ChartDataLabelsPositionAbove},
		{Name: "Sheet1!$A$31", Categories: "Sheet1!$B$29:$D$29", Values: "Sheet1!$B$31:$D$31", Sizes: "Sheet1!$B$31:$D$31", DataLabelPosition: ChartDataLabelsPositionLeft},
		{Name: "Sheet1!$A$32", Categories: "Sheet1!$B$29:$D$29", Values: "Sheet1!$B$32:$D$32", Sizes: "Sheet1!$B$32:$D$32", DataLabelPosition: ChartDataLabelsPositionBestFit},
		{Name: "Sheet1!$A$33", Categories: "Sheet1!$B$29:$D$29", Values: "Sheet1!$B$33:$D$33", Sizes: "Sheet1!$B$33:$D$33", DataLabelPosition: ChartDataLabelsPositionCenter},
		{Name: "Sheet1!$A$34", Categories: "Sheet1!$B$29:$D$29", Values: "Sheet1!$B$34:$D$34", Sizes: "Sheet1!$B$34:$D$34", DataLabelPosition: ChartDataLabelsPositionInsideBase},
		{Name: "Sheet1!$A$35", Categories: "Sheet1!$B$29:$D$29", Values: "Sheet1!$B$35:$D$35", Sizes: "Sheet1!$B$35:$D$35", DataLabelPosition: ChartDataLabelsPositionInsideEnd},
		{Name: "Sheet1!$A$36", Categories: "Sheet1!$B$29:$D$29", Values: "Sheet1!$B$36:$D$36", Sizes: "Sheet1!$B$36:$D$36", DataLabelPosition: ChartDataLabelsPositionOutsideEnd},
		{Name: "Sheet1!$A$37", Categories: "Sheet1!$B$29:$D$29", Values: "Sheet1!$B$37:$D$37", Sizes: "Sheet1!$B$37:$D$37", DataLabelPosition: ChartDataLabelsPositionRight},
	}
	format := GraphicOptions{
		ScaleX:          defaultDrawingScale,
		ScaleY:          defaultDrawingScale,
		OffsetX:         15,
		OffsetY:         10,
		PrintObject:     boolPtr(true),
		LockAspectRatio: false,
		Locked:          boolPtr(false),
	}
	legend := ChartLegend{Position: "left", ShowLegendKey: false}
	plotArea := ChartPlotArea{
		SecondPlotValues: 3,
		ShowBubbleSize:   true,
		ShowCatName:      true,
		ShowLeaderLines:  false,
		ShowPercent:      true,
		ShowSerName:      true,
		ShowVal:          true,
		Fill:             Fill{Type: "pattern", Pattern: 1},
	}
	for _, c := range []struct {
		sheetName, cell string
		opts            *Chart
	}{
		{sheetName: "Sheet1", cell: "P1", opts: &Chart{Type: Col, Series: series, Format: format, Legend: ChartLegend{Position: "none", ShowLegendKey: true}, Title: []RichTextRun{{Text: "2D Column Chart", Font: &Font{Size: 11, Family: "Calibri"}}}, PlotArea: plotArea, Border: ChartLine{Type: ChartLineNone}, ShowBlanksAs: "zero", XAxis: ChartAxis{Font: Font{Bold: true, Italic: true, Underline: "dbl", Family: "Times New Roman", Size: 15, Strike: true, Color: "000000"}, Title: []RichTextRun{{Text: "Primary Horizontal Axis Title"}}}, YAxis: ChartAxis{Font: Font{Bold: false, Italic: false, Underline: "sng", Color: "777777"}, Title: []RichTextRun{{Text: "Primary Vertical Axis Title", Font: &Font{Color: "777777", Bold: true, Italic: true, Size: 12}}}}}},
		{sheetName: "Sheet1", cell: "X1", opts: &Chart{Type: ColStacked, Series: series, Format: format, Legend: legend, Title: []RichTextRun{{Text: "2D Stacked Column Chart"}}, PlotArea: plotArea, Fill: Fill{Type: "pattern", Pattern: 1}, Border: ChartLine{Type: ChartLineAutomatic}, ShowBlanksAs: "zero", GapWidth: uintPtr(10), Overlap: intPtr(100)}},
		{sheetName: "Sheet1", cell: "P16", opts: &Chart{Type: ColPercentStacked, Series: series, Format: format, Legend: legend, Title: []RichTextRun{{Text: "100% Stacked Column Chart"}}, PlotArea: plotArea, Fill: Fill{Type: "pattern", Color: []string{"EEEEEE"}, Pattern: 1}, Border: ChartLine{Type: ChartLineSolid, Width: 2}, ShowBlanksAs: "zero", XAxis: ChartAxis{Alignment: Alignment{Vertical: "wordArtVertRtl", TextRotation: 0}}}},
		{sheetName: "Sheet1", cell: "X16", opts: &Chart{Type: Col3DClustered, Series: series, Format: format, Legend: ChartLegend{Position: "bottom", ShowLegendKey: false, Font: &Font{Size: 10}}, Title: []RichTextRun{{Text: "3D Clustered Column Chart"}}, PlotArea: plotArea, ShowBlanksAs: "zero"}},
		{sheetName: "Sheet1", cell: "P30", opts: &Chart{Type: Col3DStacked, Series: series, Format: format, Legend: legend, Title: []RichTextRun{{Text: "3D Stacked Column Chart"}}, PlotArea: plotArea, ShowBlanksAs: "zero", XAxis: ChartAxis{Alignment: Alignment{Vertical: "vert", TextRotation: 0}}}},
		{sheetName: "Sheet1", cell: "X30", opts: &Chart{Type: Col3DPercentStacked, Series: series, Format: format, Legend: legend, Title: []RichTextRun{{Text: "3D 100% Stacked Column Chart"}}, PlotArea: plotArea, ShowBlanksAs: "zero"}},
		{sheetName: "Sheet1", cell: "X45", opts: &Chart{Type: Radar, Series: series, Format: format, Legend: ChartLegend{Position: "top_right", ShowLegendKey: false}, Title: []RichTextRun{{Text: "Radar Chart"}}, PlotArea: plotArea, ShowBlanksAs: "span"}},
		{sheetName: "Sheet1", cell: "AF1", opts: &Chart{Type: Col3DConeStacked, Series: series, Format: format, Legend: legend, Title: []RichTextRun{{Text: "3D Column Cone Stacked Chart"}}, PlotArea: plotArea, ShowBlanksAs: "zero"}},
		{sheetName: "Sheet1", cell: "AF16", opts: &Chart{Type: Col3DConeClustered, Series: series, Format: format, Legend: legend, Title: []RichTextRun{{Text: "3D Column Cone Clustered Chart"}}, PlotArea: plotArea, ShowBlanksAs: "zero"}},
		{sheetName: "Sheet1", cell: "AF30", opts: &Chart{Type: Col3DConePercentStacked, Series: series, Format: format, Legend: legend, Title: []RichTextRun{{Text: "3D Column Cone Percent Stacked Chart"}}, PlotArea: plotArea, ShowBlanksAs: "zero"}},
		{sheetName: "Sheet1", cell: "AF45", opts: &Chart{Type: Col3DCone, Series: series, Format: format, Legend: legend, Title: []RichTextRun{{Text: "3D Column Cone Chart"}}, PlotArea: plotArea, ShowBlanksAs: "zero"}},
		{sheetName: "Sheet1", cell: "AN1", opts: &Chart{Type: Col3DPyramidStacked, Series: series, Format: format, Legend: legend, Title: []RichTextRun{{Text: "3D Column Pyramid Percent Stacked Chart"}}, PlotArea: plotArea, ShowBlanksAs: "zero"}},
		{sheetName: "Sheet1", cell: "AN16", opts: &Chart{Type: Col3DPyramidClustered, Series: series, Format: format, Legend: legend, Title: []RichTextRun{{Text: "3D Column Pyramid Clustered Chart"}}, PlotArea: plotArea, ShowBlanksAs: "zero"}},
		{sheetName: "Sheet1", cell: "AN30", opts: &Chart{Type: Col3DPyramidPercentStacked, Series: series, Format: format, Legend: legend, Title: []RichTextRun{{Text: "3D Column Pyramid Percent Stacked Chart"}}, PlotArea: plotArea, ShowBlanksAs: "zero"}},
		{sheetName: "Sheet1", cell: "AN45", opts: &Chart{Type: Col3DPyramid, Series: series, Format: format, Legend: legend, Title: []RichTextRun{{Text: "3D Column Pyramid Chart"}}, PlotArea: plotArea, ShowBlanksAs: "zero"}},
		{sheetName: "Sheet1", cell: "AV1", opts: &Chart{Type: Col3DCylinderStacked, Series: series, Format: format, Legend: legend, Title: []RichTextRun{{Text: "3D Column Cylinder Stacked Chart"}}, PlotArea: plotArea, ShowBlanksAs: "zero"}},
		{sheetName: "Sheet1", cell: "AV16", opts: &Chart{Type: Col3DCylinderClustered, Series: series, Format: format, Legend: legend, Title: []RichTextRun{{Text: "3D Column Cylinder Clustered Chart"}}, PlotArea: plotArea, ShowBlanksAs: "zero"}},
		{sheetName: "Sheet1", cell: "AV30", opts: &Chart{Type: Col3DCylinderPercentStacked, Series: series, Format: format, Legend: legend, Title: []RichTextRun{{Text: "3D Column Cylinder Percent Stacked Chart"}}, PlotArea: plotArea, ShowBlanksAs: "zero"}},
		{sheetName: "Sheet1", cell: "AV45", opts: &Chart{Type: Col3DCylinder, Series: series, Format: format, Legend: legend, Title: []RichTextRun{{Text: "3D Column Cylinder Chart"}}, PlotArea: plotArea, ShowBlanksAs: "zero"}},
		{sheetName: "Sheet1", cell: "P45", opts: &Chart{Type: Col3D, Series: series, Format: format, Legend: legend, Title: []RichTextRun{{Text: "3D Column Chart"}}, PlotArea: plotArea, ShowBlanksAs: "zero", XAxis: ChartAxis{Alignment: Alignment{Vertical: "vert270", TextRotation: 0}}}},
		{sheetName: "Sheet2", cell: "P1", opts: &Chart{Type: Line3D, Series: series2, Format: format, Legend: ChartLegend{Position: "top", ShowLegendKey: false}, Title: []RichTextRun{{Text: "3D Line Chart"}}, PlotArea: plotArea, ShowBlanksAs: "zero", XAxis: ChartAxis{MajorGridLines: true, MinorGridLines: true, TickLabelSkip: 1, NumFmt: ChartNumFmt{CustomNumFmt: "General"}}, YAxis: ChartAxis{MajorGridLines: true, MinorGridLines: true, MajorUnit: 1, NumFmt: ChartNumFmt{CustomNumFmt: "General"}}}},
		{sheetName: "Sheet2", cell: "X1", opts: &Chart{Type: Scatter, Series: series, Format: format, Legend: ChartLegend{Position: "bottom", ShowLegendKey: false}, Title: []RichTextRun{{Text: "Scatter Chart"}}, PlotArea: plotArea, ShowBlanksAs: "zero"}},
		{sheetName: "Sheet2", cell: "P16", opts: &Chart{Type: Doughnut, Series: series3, Format: format, Legend: ChartLegend{Position: "right", ShowLegendKey: false}, Title: []RichTextRun{{Text: "Doughnut Chart"}}, PlotArea: ChartPlotArea{ShowBubbleSize: false, ShowCatName: false, ShowLeaderLines: false, ShowPercent: true, ShowSerName: false, ShowVal: false}, ShowBlanksAs: "zero", HoleSize: 30}},
		{sheetName: "Sheet2", cell: "X16", opts: &Chart{Type: Line, Series: series2, Format: format, Legend: ChartLegend{Position: "top", ShowLegendKey: false}, Title: []RichTextRun{{Text: "Line Chart"}}, PlotArea: plotArea, ShowBlanksAs: "zero", XAxis: ChartAxis{MajorGridLines: true, MinorGridLines: true, TickLabelSkip: 1, TickLabelPosition: ChartTickLabelLow}, YAxis: ChartAxis{MajorGridLines: true, MinorGridLines: true, MajorUnit: 1}}},
		{sheetName: "Sheet2", cell: "P32", opts: &Chart{Type: Pie3D, Series: series3, Format: format, Legend: ChartLegend{Position: "bottom", ShowLegendKey: false}, Title: []RichTextRun{{Text: "3D Column Chart"}}, PlotArea: plotArea, ShowBlanksAs: "zero"}},
		{sheetName: "Sheet2", cell: "X32", opts: &Chart{Type: Pie, Series: series4, Format: format, Legend: ChartLegend{Position: "bottom", ShowLegendKey: false}, Title: []RichTextRun{{Text: "Pie Chart"}}, PlotArea: ChartPlotArea{ShowBubbleSize: true, ShowCatName: false, ShowLeaderLines: false, ShowPercent: true, ShowSerName: false, ShowVal: false, NumFmt: ChartNumFmt{CustomNumFmt: "0.00%;0;;"}}, ShowBlanksAs: "gap"}},
		// bar series chart
		{sheetName: "Sheet2", cell: "P48", opts: &Chart{Type: Bar, Series: series, Format: format, Legend: legend, Title: []RichTextRun{{Text: "2D Clustered Bar Chart"}}, PlotArea: plotArea, ShowBlanksAs: "zero"}},
		{sheetName: "Sheet2", cell: "X48", opts: &Chart{Type: BarStacked, Series: series, Format: format, Legend: legend, Title: []RichTextRun{{Text: "2D Stacked Bar Chart"}}, PlotArea: plotArea, ShowBlanksAs: "zero"}},
		{sheetName: "Sheet2", cell: "P64", opts: &Chart{Type: BarPercentStacked, Series: series, Format: format, Legend: legend, Title: []RichTextRun{{Text: "2D Stacked 100% Bar Chart"}}, PlotArea: plotArea, ShowBlanksAs: "zero"}},
		{sheetName: "Sheet2", cell: "X64", opts: &Chart{Type: Bar3DClustered, Series: series, Format: format, Legend: legend, Title: []RichTextRun{{Text: "3D Clustered Bar Chart"}}, PlotArea: plotArea, ShowBlanksAs: "zero"}},
		{sheetName: "Sheet2", cell: "P80", opts: &Chart{Type: Bar3DStacked, Series: series, Format: format, Legend: legend, Title: []RichTextRun{{Text: "3D Stacked Bar Chart"}}, PlotArea: plotArea, ShowBlanksAs: "zero", YAxis: ChartAxis{Maximum: &maximum, Minimum: &minimum}}},
		{sheetName: "Sheet2", cell: "X80", opts: &Chart{Type: Bar3DPercentStacked, Series: series, Format: format, Legend: legend, Title: []RichTextRun{{Text: "3D 100% Stacked Bar Chart"}}, PlotArea: plotArea, ShowBlanksAs: "zero", XAxis: ChartAxis{ReverseOrder: true, Secondary: true, Minimum: &zero}, YAxis: ChartAxis{ReverseOrder: true, Minimum: &zero}}},
		// area series chart
		{sheetName: "Sheet2", cell: "AF1", opts: &Chart{Type: Area, Series: series, Format: format, Legend: legend, Title: []RichTextRun{{Text: "2D Area Chart"}}, PlotArea: plotArea, ShowBlanksAs: "zero"}},
		{sheetName: "Sheet2", cell: "AN1", opts: &Chart{Type: AreaStacked, Series: series, Format: format, Legend: legend, Title: []RichTextRun{{Text: "2D Stacked Area Chart"}}, PlotArea: plotArea, ShowBlanksAs: "zero"}},
		{sheetName: "Sheet2", cell: "AF16", opts: &Chart{Type: AreaPercentStacked, Series: series, Format: format, Legend: legend, Title: []RichTextRun{{Text: "2D 100% Stacked Area Chart"}}, PlotArea: plotArea, ShowBlanksAs: "zero"}},
		{sheetName: "Sheet2", cell: "AN16", opts: &Chart{Type: Area3D, Series: series, Format: format, Legend: legend, Title: []RichTextRun{{Text: "3D Area Chart"}}, PlotArea: plotArea, ShowBlanksAs: "zero"}},
		{sheetName: "Sheet2", cell: "AF32", opts: &Chart{Type: Area3DStacked, Series: series, Format: format, Legend: legend, Title: []RichTextRun{{Text: "3D Stacked Area Chart"}}, PlotArea: plotArea, ShowBlanksAs: "zero"}},
		{sheetName: "Sheet2", cell: "AN32", opts: &Chart{Type: Area3DPercentStacked, Series: series, Format: format, Legend: legend, Title: []RichTextRun{{Text: "3D 100% Stacked Area Chart"}}, PlotArea: plotArea, ShowBlanksAs: "zero"}},
		// cylinder series chart
		{sheetName: "Sheet2", cell: "AF48", opts: &Chart{Type: Bar3DCylinderStacked, Series: series, Format: format, Legend: legend, Title: []RichTextRun{{Text: "3D Bar Cylinder Stacked Chart"}}, PlotArea: plotArea, ShowBlanksAs: "zero"}},
		{sheetName: "Sheet2", cell: "AF64", opts: &Chart{Type: Bar3DCylinderClustered, Series: series, Format: format, Legend: legend, Title: []RichTextRun{{Text: "3D Bar Cylinder Clustered Chart"}}, PlotArea: plotArea, ShowBlanksAs: "zero"}},
		{sheetName: "Sheet2", cell: "AF80", opts: &Chart{Type: Bar3DCylinderPercentStacked, Series: series, Format: format, Legend: legend, Title: []RichTextRun{{Text: "3D Bar Cylinder Percent Stacked Chart"}}, PlotArea: plotArea, ShowBlanksAs: "zero"}},
		// cone series chart
		{sheetName: "Sheet2", cell: "AN48", opts: &Chart{Type: Bar3DConeStacked, Series: series, Format: format, Legend: legend, Title: []RichTextRun{{Text: "3D Bar Cone Stacked Chart"}}, PlotArea: plotArea, ShowBlanksAs: "zero"}},
		{sheetName: "Sheet2", cell: "AN64", opts: &Chart{Type: Bar3DConeClustered, Series: series, Format: format, Legend: legend, Title: []RichTextRun{{Text: "3D Bar Cone Clustered Chart"}}, PlotArea: plotArea, ShowBlanksAs: "zero"}},
		{sheetName: "Sheet2", cell: "AN80", opts: &Chart{Type: Bar3DConePercentStacked, Series: series, Format: format, Legend: legend, Title: []RichTextRun{{Text: "3D Bar Cone Percent Stacked Chart"}}, PlotArea: plotArea, ShowBlanksAs: "zero"}},
		{sheetName: "Sheet2", cell: "AV48", opts: &Chart{Type: Bar3DPyramidStacked, Series: series, Format: format, Legend: legend, Title: []RichTextRun{{Text: "3D Bar Pyramid Stacked Chart"}}, PlotArea: plotArea, ShowBlanksAs: "zero"}},
		{sheetName: "Sheet2", cell: "AV64", opts: &Chart{Type: Bar3DPyramidClustered, Series: series, Format: format, Legend: legend, Title: []RichTextRun{{Text: "3D Bar Pyramid Clustered Chart"}}, PlotArea: plotArea, ShowBlanksAs: "zero"}},
		{sheetName: "Sheet2", cell: "AV80", opts: &Chart{Type: Bar3DPyramidPercentStacked, Series: series, Format: format, Legend: legend, Title: []RichTextRun{{Text: "3D Bar Pyramid Percent Stacked Chart"}}, PlotArea: plotArea, ShowBlanksAs: "zero"}},
		// surface series chart
		{sheetName: "Sheet2", cell: "AV1", opts: &Chart{Type: Surface3D, Series: series, Format: format, Legend: legend, Title: []RichTextRun{{Text: "3D Surface Chart"}}, PlotArea: plotArea, ShowBlanksAs: "zero", YAxis: ChartAxis{MajorGridLines: true}}},
		{sheetName: "Sheet2", cell: "AV16", opts: &Chart{Type: WireframeSurface3D, Series: series, Format: format, Legend: legend, Title: []RichTextRun{{Text: "3D Wireframe Surface Chart"}}, PlotArea: plotArea, ShowBlanksAs: "zero", YAxis: ChartAxis{MajorGridLines: true}}},
		{sheetName: "Sheet2", cell: "AV32", opts: &Chart{Type: Contour, Series: series, Format: format, Legend: legend, Title: []RichTextRun{{Text: "Contour Chart"}}, PlotArea: plotArea, ShowBlanksAs: "zero"}},
		{sheetName: "Sheet2", cell: "BD1", opts: &Chart{Type: WireframeContour, Series: series, Format: format, Legend: legend, Title: []RichTextRun{{Text: "Wireframe Contour Chart"}}, PlotArea: plotArea, ShowBlanksAs: "zero"}},
		// bubble chart
		{sheetName: "Sheet2", cell: "BD16", opts: &Chart{Type: Bubble, Series: series5, Format: format, Legend: legend, Title: []RichTextRun{{Text: "Bubble Chart"}}, PlotArea: plotArea, ShowBlanksAs: "zero", BubbleSize: 75}},
		{sheetName: "Sheet2", cell: "BD32", opts: &Chart{Type: Bubble3D, Series: series5, Format: format, Legend: legend, Title: []RichTextRun{{Text: "Bubble 3D Chart"}}, PlotArea: plotArea, ShowBlanksAs: "zero", XAxis: ChartAxis{MajorGridLines: true}, YAxis: ChartAxis{MajorGridLines: true}}},
		// pie of pie chart
		{sheetName: "Sheet2", cell: "BD48", opts: &Chart{Type: PieOfPie, Series: series3, Format: format, Legend: legend, Title: []RichTextRun{{Text: "Pie of Pie Chart"}}, PlotArea: plotArea, ShowBlanksAs: "zero", XAxis: ChartAxis{MajorGridLines: true}, YAxis: ChartAxis{MajorGridLines: true}}},
		// bar of pie chart
		{sheetName: "Sheet2", cell: "BD64", opts: &Chart{Type: BarOfPie, Series: series3, Format: format, Legend: legend, Title: []RichTextRun{{Text: "Bar of Pie Chart"}}, PlotArea: plotArea, ShowBlanksAs: "zero", XAxis: ChartAxis{MajorGridLines: true}, YAxis: ChartAxis{MajorGridLines: true}}},
	} {
		assert.NoError(t, f.AddChart(c.sheetName, c.cell, c.opts))
	}
	// combo chart
	_, err = f.NewSheet("Combo Charts")
	assert.NoError(t, err)
	clusteredColumnCombo := [][]interface{}{
		{"A1", Line, "Clustered Column - Line Chart"},
		{"I1", Doughnut, "Clustered Column - Doughnut Chart"},
	}
	for _, props := range clusteredColumnCombo {
		assert.NoError(t, f.AddChart("Combo Charts", props[0].(string), &Chart{Type: Col, Series: series[:4], Format: format, Legend: legend, Title: []RichTextRun{{Text: props[2].(string)}}, PlotArea: ChartPlotArea{ShowBubbleSize: true, ShowCatName: false, ShowDataTable: true, ShowDataTableKeys: true, ShowLeaderLines: false, ShowPercent: true, ShowSerName: true, ShowVal: true}}, &Chart{Type: props[1].(ChartType), Series: series[4:], Format: format, Legend: legend, PlotArea: ChartPlotArea{ShowBubbleSize: true, ShowCatName: false, ShowLeaderLines: false, ShowPercent: true, ShowSerName: true, ShowVal: true}, YAxis: ChartAxis{Secondary: true}}))
	}
	stackedAreaCombo := map[string][]interface{}{
		"A16": {Line, "Stacked Area - Line Chart"},
		"I16": {Doughnut, "Stacked Area - Doughnut Chart"},
	}
	for axis, props := range stackedAreaCombo {
		assert.NoError(t, f.AddChart("Combo Charts", axis, &Chart{Type: AreaStacked, Series: series[:4], Format: format, Legend: legend, Title: []RichTextRun{{Text: props[1].(string)}}, PlotArea: ChartPlotArea{ShowBubbleSize: true, ShowCatName: false, ShowLeaderLines: false, ShowPercent: true, ShowSerName: true, ShowVal: true}}, &Chart{Type: props[0].(ChartType), Series: series[4:], Format: format, Legend: legend, PlotArea: ChartPlotArea{ShowBubbleSize: true, ShowCatName: false, ShowLeaderLines: false, ShowPercent: true, ShowSerName: true, ShowVal: true}}))
	}
	assert.NoError(t, f.SaveAs(filepath.Join("test", "TestAddChart.xlsx")))
	// Test with invalid sheet name
	assert.EqualError(t, f.AddChart("Sheet:1", "A1", &Chart{Type: Col, Series: series[:1]}), ErrSheetNameInvalid.Error())
	// Test with illegal cell reference
	assert.EqualError(t, f.AddChart("Sheet2", "A", &Chart{Type: Col, Series: series, Format: format, Legend: legend, Title: []RichTextRun{{Text: "2D Column Chart"}}, PlotArea: plotArea, ShowBlanksAs: "zero"}), newCellNameToCoordinatesError("A", newInvalidCellNameError("A")).Error())
	// Test with unsupported chart type
	assert.EqualError(t, f.AddChart("Sheet2", "BD32", &Chart{Type: 0x39, Series: series, Format: format, Legend: legend, Title: []RichTextRun{{Text: "Bubble 3D Chart"}}, PlotArea: plotArea, ShowBlanksAs: "zero"}), newUnsupportedChartType(0x39).Error())
	// Test add combo chart with invalid format set
	assert.EqualError(t, f.AddChart("Sheet2", "BD32", &Chart{Type: Col, Series: series, Format: format, Legend: legend, Title: []RichTextRun{{Text: "2D Column Chart"}}, PlotArea: plotArea, ShowBlanksAs: "zero"}, nil), ErrParameterInvalid.Error())
	// Test add combo chart with unsupported chart type
	assert.EqualError(t, f.AddChart("Sheet2", "BD64", &Chart{Type: BarOfPie, Series: []ChartSeries{{Name: "Sheet1!$A$30", Categories: "Sheet1!$A$30:$D$37", Values: "Sheet1!$B$30:$B$37"}}, Format: format, Legend: legend, Title: []RichTextRun{{Text: "Bar of Pie Chart"}}, PlotArea: plotArea, ShowBlanksAs: "zero", XAxis: ChartAxis{MajorGridLines: true}, YAxis: ChartAxis{MajorGridLines: true}}, &Chart{Type: 0x39, Series: []ChartSeries{{Name: "Sheet1!$A$30", Categories: "Sheet1!$A$30:$D$37", Values: "Sheet1!$B$30:$B$37"}}, Format: format, Legend: legend, Title: []RichTextRun{{Text: "Bar of Pie Chart"}}, PlotArea: plotArea, ShowBlanksAs: "zero", XAxis: ChartAxis{MajorGridLines: true}, YAxis: ChartAxis{MajorGridLines: true}}), newUnsupportedChartType(0x39).Error())
	// Test add chart with series transparency value exceeds limit
	assert.Equal(t, ErrTransparency, f.AddChart("Sheet1", "BD64", &Chart{Type: Col, Series: []ChartSeries{{Name: "Sheet1!$A$30", Categories: "Sheet1!$B$29:$D$29", Values: "Sheet1!$B$30:$D$30", Fill: Fill{Transparency: 110}}}}))
	// Test add chart with transparency value exceeds limit
	assert.Equal(t, ErrTransparency, f.AddChart("Sheet1", "BD64", &Chart{Type: Col, Series: []ChartSeries{{Name: "Sheet1!$A$30", Categories: "Sheet1!$B$29:$D$29", Values: "Sheet1!$B$30:$D$30"}}, Fill: Fill{Transparency: -1}}))
	assert.NoError(t, f.Close())

	// Test add chart with unsupported charset content types.
	f.ContentTypes = nil
	f.Pkg.Store(defaultXMLPathContentTypes, MacintoshCyrillicCharset)
	assert.EqualError(t, f.AddChart("Sheet1", "P1", &Chart{Type: Col, Series: []ChartSeries{{Name: "Sheet1!$A$30", Categories: "Sheet1!$B$29:$D$29", Values: "Sheet1!$B$30:$D$30"}}, Title: []RichTextRun{{Text: "2D Column Chart"}}}), "XML syntax error on line 1: invalid UTF-8")

	t.Run("for_create_stock_chart", func(t *testing.T) {
		f := NewFile()
		for i, row := range [][]interface{}{
			{"Date", "Volume", "Open", "High", "Low", "Close"},
			{45593, 14864000, 431.66, 431.94, 426.3, 426.59},
			{45590, 16899100, 426.76, 432.52, 426.57, 428.15},
			{45589, 13581600, 425.33, 425.98, 422.4, 424.73},
			{45588, 19654400, 430.86, 431.08, 422.53, 424.6},
			{45587, 25482200, 418.49, 430.58, 418.04, 427.51},
			{45586, 14206100, 416.12, 418.96, 413.75, 418.78},
			{45583, 17145300, 417.14, 419.65, 416.26, 418.16},
			{45582, 14820000, 422.36, 422.5, 415.59, 416.72},
			{45581, 15508900, 415.17, 416.36, 410.48, 416.12},
			{45580, 18900200, 422.18, 422.48, 415.26, 418.74},
			{45579, 16653100, 417.77, 424.04, 417.52, 419.14},
			{45576, 14144900, 416.14, 417.13, 413.25, 416.32},
			{45575, 13848400, 415.23, 417.35, 413.15, 415.84},
			{45574, 14974300, 415.86, 420.38, 414.3, 417.46},
			{45573, 19229300, 410.9, 415.66, 408.17, 414.71},
			{45572, 20919800, 416, 417.11, 409, 409.54},
			{45569, 19169700, 418.24, 419.75, 414.97, 416.06},
			{45568, 13686400, 417.63, 419.55, 414.29, 416.54},
			{45567, 16582300, 422.58, 422.82, 416.71, 417.13},
			{45566, 19092900, 428.45, 428.48, 418.81, 420.69},
			{45565, 16807300, 428.21, 430.42, 425.37, 430.3},
		} {
			cell, err := CoordinatesToCellName(1, i+1)
			assert.NoError(t, err)
			assert.NoError(t, f.SetSheetRow("Sheet1", cell, &row))
		}
		style, err := f.NewStyle(&Style{NumFmt: 15})
		assert.NoError(t, err)
		assert.NoError(t, f.SetColStyle("Sheet1", "A", style))

		assert.NoError(t, f.AddChart("Sheet1", "G1", &Chart{
			Type: StockHighLowClose,
			Series: []ChartSeries{
				{Name: "Sheet1!$D$1", Categories: "Sheet1!$A$2:$A$22", Values: "Sheet1!$D$2:$D$22"},
				{Name: "Sheet1!$E$1", Categories: "Sheet1!$A$2:$A$22", Values: "Sheet1!$E$2:$E$22"},
				{Name: "Sheet1!$F$1", Categories: "Sheet1!$A$2:$A$22", Values: "Sheet1!$F$2:$F$22"},
			},
			Legend: ChartLegend{Position: "none"},
			Title:  []RichTextRun{{Text: "High-Low-Close Stock Chart"}},
			XAxis:  ChartAxis{NumFmt: ChartNumFmt{CustomNumFmt: "d-mmm-yy"}},
		}))
		assert.NoError(t, f.AddChart("Sheet1", "G16", &Chart{
			Type: StockOpenHighLowClose,
			Series: []ChartSeries{
				{Name: "Sheet1!$C$1", Categories: "Sheet1!$A$2:$A$22", Values: "Sheet1!$C$2:$C$22"},
				{Name: "Sheet1!$D$1", Categories: "Sheet1!$A$2:$A$22", Values: "Sheet1!$D$2:$D$22"},
				{Name: "Sheet1!$E$1", Categories: "Sheet1!$A$2:$A$22", Values: "Sheet1!$E$2:$E$22"},
				{Name: "Sheet1!$F$1", Categories: "Sheet1!$A$2:$A$22", Values: "Sheet1!$F$2:$F$22"},
			},
			Legend: ChartLegend{Position: "none"},
			Title:  []RichTextRun{{Text: "Open-High-Low-Close Stock Chart"}},
			XAxis:  ChartAxis{NumFmt: ChartNumFmt{CustomNumFmt: "d-mmm-yy"}},
			PlotArea: ChartPlotArea{
				UpBars: ChartUpDownBar{
					Border: ChartLine{Type: ChartLineNone},
					Fill:   Fill{Type: "pattern", Color: []string{"00B050"}, Pattern: 1},
				},
				DownBars: ChartUpDownBar{
					Border: ChartLine{Type: ChartLineNone},
					Fill:   Fill{Type: "pattern", Color: []string{"FF0000"}, Pattern: 1},
				},
			},
		}))
		assert.NoError(t, f.AddChart("Sheet1", "O1", &Chart{
			Type: Col,
			Series: []ChartSeries{
				{Name: "Sheet1!$B$1", Categories: "Sheet1!$A$2:$A$22", Values: "Sheet1!$B$2:$B$22"},
			},
			VaryColors: boolPtr(false),
			XAxis:      ChartAxis{NumFmt: ChartNumFmt{CustomNumFmt: "d-mmm-yy"}},
			YAxis:      ChartAxis{NumFmt: ChartNumFmt{CustomNumFmt: "#,##0"}},
			Title:      []RichTextRun{{Text: "Volume-High-Low-Close Stock Chart"}},
		}, &Chart{
			Type: StockHighLowClose,
			Series: []ChartSeries{
				{Name: "Sheet1!$D$1", Categories: "Sheet1!$A$2:$A$22", Values: "Sheet1!$D$2:$D$22"},
				{Name: "Sheet1!$E$1", Categories: "Sheet1!$A$2:$A$22", Values: "Sheet1!$E$2:$E$22"},
				{Name: "Sheet1!$F$1", Categories: "Sheet1!$A$2:$A$22", Values: "Sheet1!$F$2:$F$22"},
			},
			YAxis: ChartAxis{Secondary: true},
		}))
		assert.NoError(t, f.AddChart("Sheet1", "O16", &Chart{
			Type: Col,
			Series: []ChartSeries{
				{Name: "Sheet1!$B$1", Categories: "Sheet1!$A$2:$A$22", Values: "Sheet1!$B$2:$B$22"},
			},
			VaryColors: boolPtr(false),
			XAxis:      ChartAxis{NumFmt: ChartNumFmt{CustomNumFmt: "d-mmm-yy"}},
			YAxis:      ChartAxis{NumFmt: ChartNumFmt{CustomNumFmt: "#,##0"}},
			Title:      []RichTextRun{{Text: "Volume-Open-High-Low-Close Stock Chart"}},
		}, &Chart{
			Type: StockOpenHighLowClose,
			Series: []ChartSeries{
				{Name: "Sheet1!$C$1", Categories: "Sheet1!$A$2:$A$22", Values: "Sheet1!$C$2:$C$22"},
				{Name: "Sheet1!$D$1", Categories: "Sheet1!$A$2:$A$22", Values: "Sheet1!$D$2:$D$22"},
				{Name: "Sheet1!$E$1", Categories: "Sheet1!$A$2:$A$22", Values: "Sheet1!$E$2:$E$22"},
				{Name: "Sheet1!$F$1", Categories: "Sheet1!$A$2:$A$22", Values: "Sheet1!$F$2:$F$22"},
			},
			YAxis: ChartAxis{Secondary: true},
		}))
		assert.NoError(t, f.SaveAs(filepath.Join("test", "TestAddChartStock.xlsx")))
	})
}

func TestAddChartSheet(t *testing.T) {
	categories := map[string]string{"A2": "Small", "A3": "Normal", "A4": "Large", "B1": "Apple", "C1": "Orange", "D1": "Pear"}
	values := map[string]int{"B2": 2, "C2": 3, "D2": 3, "B3": 5, "C3": 2, "D3": 4, "B4": 6, "C4": 7, "D4": 8}
	f := NewFile()
	for k, v := range categories {
		assert.NoError(t, f.SetCellValue("Sheet1", k, v))
	}
	for k, v := range values {
		assert.NoError(t, f.SetCellValue("Sheet1", k, v))
	}
	series := []ChartSeries{
		{Name: "Sheet1!$A$2", Categories: "Sheet1!$B$1:$D$1", Values: "Sheet1!$B$2:$D$2"},
		{Name: "Sheet1!$A$3", Categories: "Sheet1!$B$1:$D$1", Values: "Sheet1!$B$3:$D$3"},
		{Name: "Sheet1!$A$4", Categories: "Sheet1!$B$1:$D$1", Values: "Sheet1!$B$4:$D$4"},
	}
	assert.NoError(t, f.AddChartSheet("Chart1", &Chart{Type: Col3DClustered, Series: series, Title: []RichTextRun{{Text: "Fruit 3D Clustered Column Chart"}}}))
	// Test set the chartsheet as active sheet
	var sheetIdx int
	for idx, sheetName := range f.GetSheetList() {
		if sheetName != "Chart1" {
			continue
		}
		sheetIdx = idx
	}
	f.SetActiveSheet(sheetIdx)

	// Test cell value on chartsheet
	assert.EqualError(t, f.SetCellValue("Chart1", "A1", true), "sheet Chart1 is not a worksheet")
	// Test add chartsheet on already existing name sheet

	assert.EqualError(t, f.AddChartSheet("Sheet1", &Chart{Type: Col3DClustered, Series: series, Title: []RichTextRun{{Text: "Fruit 3D Clustered Column Chart"}}}), ErrExistsSheet.Error())
	// Test add chartsheet with invalid sheet name
	assert.EqualError(t, f.AddChartSheet("Sheet:1", nil, &Chart{Type: Col3DClustered, Series: series, Title: []RichTextRun{{Text: "Fruit 3D Clustered Column Chart"}}}), ErrSheetNameInvalid.Error())
	// Test with unsupported chart type
	assert.EqualError(t, f.AddChartSheet("Chart2", &Chart{Type: 0x39, Series: series, Title: []RichTextRun{{Text: "Fruit 3D Clustered Column Chart"}}}), newUnsupportedChartType(0x39).Error())

	assert.NoError(t, f.UpdateLinkedValue())

	assert.NoError(t, f.SaveAs(filepath.Join("test", "TestAddChartSheet.xlsx")))
	// Test add chart sheet with unsupported charset content types
	f = NewFile()
	f.ContentTypes = nil
	f.Pkg.Store(defaultXMLPathContentTypes, MacintoshCyrillicCharset)
	assert.EqualError(t, f.AddChartSheet("Chart4", &Chart{Type: Col, Series: []ChartSeries{{Name: "Sheet1!$A$30", Categories: "Sheet1!$B$29:$D$29", Values: "Sheet1!$B$30:$D$30"}}, Title: []RichTextRun{{Text: "2D Column Chart"}}}), "XML syntax error on line 1: invalid UTF-8")
}

func TestDeleteChart(t *testing.T) {
	f, err := OpenFile(filepath.Join("test", "Book1.xlsx"))
	assert.NoError(t, err)
	assert.NoError(t, f.DeleteChart("Sheet1", "A1"))
	series := []ChartSeries{
		{Name: "Sheet1!$A$30", Categories: "Sheet1!$B$29:$D$29", Values: "Sheet1!$B$30:$D$30"},
		{Name: "Sheet1!$A$31", Categories: "Sheet1!$B$29:$D$29", Values: "Sheet1!$B$31:$D$31"},
		{Name: "Sheet1!$A$32", Categories: "Sheet1!$B$29:$D$29", Values: "Sheet1!$B$32:$D$32"},
		{Name: "Sheet1!$A$33", Categories: "Sheet1!$B$29:$D$29", Values: "Sheet1!$B$33:$D$33"},
		{Name: "Sheet1!$A$34", Categories: "Sheet1!$B$29:$D$29", Values: "Sheet1!$B$34:$D$34"},
		{Name: "Sheet1!$A$35", Categories: "Sheet1!$B$29:$D$29", Values: "Sheet1!$B$35:$D$35"},
		{Name: "Sheet1!$A$36", Categories: "Sheet1!$B$29:$D$29", Values: "Sheet1!$B$36:$D$36"},
		{Name: "Sheet1!$A$37", Categories: "Sheet1!$B$29:$D$29", Values: "Sheet1!$B$37:$D$37"},
	}
	format := GraphicOptions{
		ScaleX:          defaultDrawingScale,
		ScaleY:          defaultDrawingScale,
		OffsetX:         15,
		OffsetY:         10,
		PrintObject:     boolPtr(true),
		LockAspectRatio: false,
		Locked:          boolPtr(false),
	}
	legend := ChartLegend{Position: "left", ShowLegendKey: false}
	plotArea := ChartPlotArea{
		ShowBubbleSize:  true,
		ShowCatName:     true,
		ShowLeaderLines: false,
		ShowPercent:     true,
		ShowSerName:     true,
		ShowVal:         true,
	}
	assert.NoError(t, f.AddChart("Sheet1", "P1", &Chart{Type: Col, Series: series, Format: format, Legend: legend, Title: []RichTextRun{{Text: "2D Column Chart"}}, PlotArea: plotArea, ShowBlanksAs: "zero"}))
	assert.NoError(t, f.DeleteChart("Sheet1", "P1"))
	assert.NoError(t, f.SaveAs(filepath.Join("test", "TestDeleteChart.xlsx")))
	// Test delete chart with invalid sheet name
	assert.EqualError(t, f.DeleteChart("Sheet:1", "P1"), ErrSheetNameInvalid.Error())
	// Test delete chart on not exists worksheet
	assert.EqualError(t, f.DeleteChart("SheetN", "A1"), "sheet SheetN does not exist")
	// Test delete chart with invalid coordinates
	assert.EqualError(t, f.DeleteChart("Sheet1", ""), newCellNameToCoordinatesError("", newInvalidCellNameError("")).Error())
	// Test delete chart on no chart worksheet
	assert.NoError(t, NewFile().DeleteChart("Sheet1", "A1"))
	assert.NoError(t, f.Close())
}

func TestChartWithLogarithmicBase(t *testing.T) {
	// Create test workbook with data
	f := NewFile()
	sheet1 := f.GetSheetName(0)
	categories := map[string]float64{
		"A1":  1,
		"A2":  2,
		"A3":  3,
		"A4":  4,
		"A5":  5,
		"A6":  6,
		"A7":  7,
		"A8":  8,
		"A9":  9,
		"A10": 10,
		"B1":  0.1,
		"B2":  1,
		"B3":  2,
		"B4":  3,
		"B5":  20,
		"B6":  30,
		"B7":  100,
		"B8":  500,
		"B9":  700,
		"B10": 5000,
	}
	for cell, v := range categories {
		assert.NoError(t, f.SetCellValue(sheet1, cell, v))
	}
	series := []ChartSeries{{Name: "value", Categories: "Sheet1!$A$1:$A$19", Values: "Sheet1!$B$1:$B$10"}}
	dimension := []uint{640, 480, 320, 240}
	for _, c := range []struct {
		cell string
		opts *Chart
	}{
		{cell: "C1", opts: &Chart{Type: Line, Dimension: ChartDimension{Width: dimension[0], Height: dimension[1]}, Series: series, Title: []RichTextRun{{Text: "Line chart without log scaling"}}}},
		{cell: "M1", opts: &Chart{Type: Line, Dimension: ChartDimension{Width: dimension[0], Height: dimension[1]}, Series: series, Title: []RichTextRun{{Text: "Line chart with log 10.5 scaling"}}, YAxis: ChartAxis{LogBase: 10.5}}},
		{cell: "A25", opts: &Chart{Type: Line, Dimension: ChartDimension{Width: dimension[2], Height: dimension[3]}, Series: series, Title: []RichTextRun{{Text: "Line chart with log 1.9 scaling"}}, YAxis: ChartAxis{LogBase: 1.9}}},
		{cell: "F25", opts: &Chart{Type: Line, Dimension: ChartDimension{Width: dimension[2], Height: dimension[3]}, Series: series, Title: []RichTextRun{{Text: "Line chart with log 2 scaling"}}, YAxis: ChartAxis{LogBase: 2}}},
		{cell: "K25", opts: &Chart{Type: Line, Dimension: ChartDimension{Width: dimension[2], Height: dimension[3]}, Series: series, Title: []RichTextRun{{Text: "Line chart with log 1000.1 scaling"}}, YAxis: ChartAxis{LogBase: 1000.1}}},
		{cell: "P25", opts: &Chart{Type: Line, Dimension: ChartDimension{Width: dimension[2], Height: dimension[3]}, Series: series, Title: []RichTextRun{{Text: "Line chart with log 1000 scaling"}}, YAxis: ChartAxis{LogBase: 1000}}},
	} {
		// Add two chart, one without and one with log scaling
		assert.NoError(t, f.AddChart(sheet1, c.cell, c.opts))
	}

	// Export workbook for human confirmation
	assert.NoError(t, f.SaveAs(filepath.Join("test", "TestChartWithLogarithmicBase10.xlsx")))

	// Write the workbook to a buffer
	var buffer bytes.Buffer
	assert.NoError(t, f.Write(&buffer))

	// Read back the workbook from the buffer
	newFile, err := OpenReader(&buffer)
	assert.NoError(t, err)

	// Check the number of charts
	expectedChartsCount := 6
	chartsNum := newFile.countCharts()
	if !assert.Equal(t, expectedChartsCount, chartsNum,
		"Expected %d charts, actual %d", expectedChartsCount, chartsNum) {
		t.FailNow()
	}

	chartSpaces := make([]xlsxChartSpace, expectedChartsCount)
	type xmlChartContent []byte
	xmlCharts := make([]xmlChartContent, expectedChartsCount)
	expectedChartsLogBase := []float64{0, 10.5, 0, 2, 0, 1000}
	var (
		drawingML interface{}
		ok        bool
	)
	for i := 0; i < expectedChartsCount; i++ {
		chartPath := fmt.Sprintf("xl/charts/chart%d.xml", i+1)
		if drawingML, ok = newFile.Pkg.Load(chartPath); ok {
			xmlCharts[i] = drawingML.([]byte)
		}
		assert.True(t, ok, "Can't open the %s", chartPath)

		err = xml.Unmarshal(xmlCharts[i], &chartSpaces[i])
		if !assert.NoError(t, err) {
			t.FailNow()
		}

		chartLogBasePtr := chartSpaces[i].Chart.PlotArea.ValAx[0].Scaling.LogBase
		if expectedChartsLogBase[i] == 0 {
			if !assert.Nil(t, chartLogBasePtr, "LogBase is not nil") {
				t.FailNow()
			}
		} else {
			if !assert.NotNil(t, chartLogBasePtr, "LogBase is nil") {
				t.FailNow()
			}
			if !assert.Equal(t, expectedChartsLogBase[i], *(chartLogBasePtr.Val),
				"Expected log base to %f, actual %f", expectedChartsLogBase[i], *(chartLogBasePtr.Val)) {
				t.FailNow()
			}
		}
	}
}

func TestGetChartCells(t *testing.T) {
	f := NewFile()
	sheet := f.GetSheetName(0)

	for cell, v := range map[string]interface{}{
		"A1": nil, "B1": "Apple", "C1": "Orange", "D1": "Pear",
		"A2": "Small", "B2": 2, "C2": 3, "D2": 3,
		"A3": "Normal", "B3": 5, "C3": 2, "D3": 4,
		"A4": "Large", "B4": 6, "C4": 7, "D4": 8,
	} {
		assert.NoError(t, f.SetCellValue(sheet, cell, v))
	}

	series := []ChartSeries{
		{Name: "Sheet1!$A$2", Categories: "Sheet1!$B$1:$D$1", Values: "Sheet1!$B$2:$D$2"},
		{Name: "Sheet1!$A$3", Categories: "Sheet1!$B$1:$D$1", Values: "Sheet1!$B$3:$D$3"},
	}

	// No chart on the sheet yet
	cells, err := f.GetChartCells(sheet)
	assert.NoError(t, err)
	assert.Empty(t, cells)

	// Add charts
	assert.NoError(t, f.AddChart(sheet, "E1", &Chart{Type: Col, Series: series}))
	assert.NoError(t, f.AddChart(sheet, "M1", &Chart{Type: Line, Series: series}))
	assert.NoError(t, f.AddChart(sheet, "E16", &Chart{Type: Pie, Series: series}))

	cells, err = f.GetChartCells(sheet)
	assert.NoError(t, err)
	assert.Len(t, cells, 3)
	assert.Contains(t, cells, "E1")
	assert.Contains(t, cells, "M1")
	assert.Contains(t, cells, "E16")

	// Verify round-trip: GetChartCells + GetChart
	for _, cell := range cells {
		charts, err := f.GetChart(sheet, cell)
		assert.NoError(t, err)
		assert.NotNil(t, charts)
		assert.GreaterOrEqual(t, len(charts), 1)
	}

	// Error: invalid sheet name
	_, err = f.GetChartCells("Sheet:1")
	assert.EqualError(t, err, ErrSheetNameInvalid.Error())

	// Error: non-existent sheet
	_, err = f.GetChartCells("SheetN")
	assert.EqualError(t, err, "sheet SheetN does not exist")

	// No drawing on a fresh sheet
	f2 := NewFile()
	cells, err = f2.GetChartCells("Sheet1")
	assert.NoError(t, err)
	assert.Nil(t, cells)
	assert.NoError(t, f2.Close())

	assert.NoError(t, f.Close())
}

func TestGetChart(t *testing.T) {
	f := NewFile()
	sheet := f.GetSheetName(0)

	for cell, v := range map[string]interface{}{
		"A1": nil, "B1": "Apple", "C1": "Orange", "D1": "Pear",
		"A2": "Small", "B2": 2, "C2": 3, "D2": 3,
		"A3": "Normal", "B3": 5, "C3": 2, "D3": 4,
		"A4": "Large", "B4": 6, "C4": 7, "D4": 8,
	} {
		assert.NoError(t, f.SetCellValue(sheet, cell, v))
	}

	series := []ChartSeries{
		{Name: "Sheet1!$A$2", Categories: "Sheet1!$B$1:$D$1", Values: "Sheet1!$B$2:$D$2"},
		{Name: "Sheet1!$A$3", Categories: "Sheet1!$B$1:$D$1", Values: "Sheet1!$B$3:$D$3"},
		{Name: "Sheet1!$A$4", Categories: "Sheet1!$B$1:$D$1", Values: "Sheet1!$B$4:$D$4"},
	}

	// Test get chart with various types
	for _, tc := range []struct {
		cell      string
		chartType ChartType
		title     string
	}{
		{"E1", Col, "Column Chart"},
		{"M1", Bar, "Bar Chart"},
		{"E16", Line, "Line Chart"},
		{"M16", Pie, "Pie Chart"},
		{"E32", Area, "Area Chart"},
		{"M32", Radar, "Radar Chart"},
		{"E48", Scatter, "Scatter Chart"},
		{"M48", Doughnut, "Doughnut Chart"},
	} {
		assert.NoError(t, f.AddChart(sheet, tc.cell, &Chart{
			Type:   tc.chartType,
			Series: series,
			Title:  []RichTextRun{{Text: tc.title}},
		}))
		charts, err := f.GetChart(sheet, tc.cell)
		assert.NoError(t, err)
		assert.NotNil(t, charts, tc.title)
		assert.Len(t, charts, 1, tc.title)
		assert.Equal(t, tc.chartType, charts[0].Type, tc.title)
		assert.Len(t, charts[0].Series, 3, tc.title)
		assert.Equal(t, "Sheet1!$A$2", charts[0].Series[0].Name, tc.title)
		assert.Equal(t, "Sheet1!$B$1:$D$1", charts[0].Series[0].Categories, tc.title)
		assert.Equal(t, "Sheet1!$B$2:$D$2", charts[0].Series[0].Values, tc.title)
		assert.Len(t, charts[0].Title, 1, tc.title)
		assert.Equal(t, tc.title, charts[0].Title[0].Text, tc.title)
	}

	// Test get chart with 3D types
	for _, tc := range []struct {
		cell      string
		chartType ChartType
	}{
		{"A64", Col3DClustered},
		{"J64", Bar3DClustered},
		{"A80", Pie3D},
		{"J80", Area3D},
		{"A96", Line3D},
	} {
		assert.NoError(t, f.AddChart(sheet, tc.cell, &Chart{
			Type:   tc.chartType,
			Series: series,
			Title:  []RichTextRun{{Text: "3D Chart"}},
		}))
		charts, err := f.GetChart(sheet, tc.cell)
		assert.NoError(t, err)
		assert.NotNil(t, charts)
		assert.Len(t, charts, 1)
		assert.Equal(t, tc.chartType, charts[0].Type)
	}

	// Test get chart with 3D bar shapes
	for _, tc := range []struct {
		cell      string
		chartType ChartType
	}{
		{"A112", Bar3DConeClustered},
		{"J112", Bar3DPyramidClustered},
		{"A128", Bar3DCylinderClustered},
		{"J128", Col3DCone},
		{"A144", Col3DPyramid},
		{"J144", Col3DCylinder},
	} {
		assert.NoError(t, f.AddChart(sheet, tc.cell, &Chart{
			Type:   tc.chartType,
			Series: series,
		}))
		charts, err := f.GetChart(sheet, tc.cell)
		assert.NoError(t, err)
		assert.NotNil(t, charts)
		assert.Len(t, charts, 1)
		assert.Equal(t, tc.chartType, charts[0].Type)
	}

	// Test get combo chart
	assert.NoError(t, f.AddChart(sheet, "A160", &Chart{
		Type:   Col,
		Series: series,
		Title:  []RichTextRun{{Text: "Combo Chart"}},
	}, &Chart{
		Type:   Line,
		Series: series,
	}))
	charts, err := f.GetChart(sheet, "A160")
	assert.NoError(t, err)
	assert.NotNil(t, charts)
	assert.Len(t, charts, 2)
	assert.Equal(t, Col, charts[0].Type)
	assert.Equal(t, Line, charts[1].Type)
	assert.Len(t, charts[0].Title, 1)
	assert.Equal(t, "Combo Chart", charts[0].Title[0].Text)
	assert.Nil(t, charts[1].Title)

	// Test get chart with plot area options
	assert.NoError(t, f.AddChart(sheet, "A176", &Chart{
		Type:   Col,
		Series: series,
		PlotArea: ChartPlotArea{
			ShowVal:     true,
			ShowCatName: true,
			ShowSerName: true,
		},
	}))
	charts, err = f.GetChart(sheet, "A176")
	assert.NoError(t, err)
	assert.NotNil(t, charts)
	assert.True(t, charts[0].PlotArea.ShowVal)
	assert.True(t, charts[0].PlotArea.ShowCatName)
	assert.True(t, charts[0].PlotArea.ShowSerName)

	// Test get chart on no chart cell
	charts, err = f.GetChart(sheet, "Z1")
	assert.NoError(t, err)
	assert.Nil(t, charts)

	// Test get chart on no drawing worksheet
	charts, err = f.GetChart(sheet, "Z1")
	assert.NoError(t, err)
	assert.Nil(t, charts)

	// Test get chart with Surface and Contour types
	assert.NoError(t, f.AddChart(sheet, "A192", &Chart{
		Type:   Surface3D,
		Series: series,
	}))
	charts, err = f.GetChart(sheet, "A192")
	assert.NoError(t, err)
	assert.NotNil(t, charts)
	assert.Equal(t, Surface3D, charts[0].Type)

	assert.NoError(t, f.AddChart(sheet, "J192", &Chart{
		Type:   WireframeSurface3D,
		Series: series,
	}))
	charts, err = f.GetChart(sheet, "J192")
	assert.NoError(t, err)
	assert.NotNil(t, charts)
	assert.Equal(t, WireframeSurface3D, charts[0].Type)

	// Test get chart with legend position
	assert.NoError(t, f.AddChart(sheet, "A208", &Chart{
		Type:   Col,
		Series: series,
		Legend: ChartLegend{Position: "top"},
	}))
	charts, err = f.GetChart(sheet, "A208")
	assert.NoError(t, err)
	assert.NotNil(t, charts)
	assert.Equal(t, "top", charts[0].Legend.Position)

	// Test get chart with no legend
	assert.NoError(t, f.AddChart(sheet, "J208", &Chart{
		Type:   Col,
		Series: series,
		Legend: ChartLegend{Position: "none"},
	}))
	charts, err = f.GetChart(sheet, "J208")
	assert.NoError(t, err)
	assert.NotNil(t, charts)
	assert.Equal(t, "none", charts[0].Legend.Position)

	// Test get chart with stacked and percent stacked types
	for _, tc := range []struct {
		cell      string
		chartType ChartType
	}{
		{"A224", BarStacked},
		{"J224", BarPercentStacked},
		{"A240", ColStacked},
		{"J240", ColPercentStacked},
		{"A256", AreaStacked},
		{"J256", AreaPercentStacked},
	} {
		assert.NoError(t, f.AddChart(sheet, tc.cell, &Chart{
			Type:   tc.chartType,
			Series: series,
		}))
		charts, err = f.GetChart(sheet, tc.cell)
		assert.NoError(t, err)
		assert.NotNil(t, charts)
		assert.Equal(t, tc.chartType, charts[0].Type)
	}

	// Test error cases
	// Invalid sheet name
	_, err = f.GetChart("Sheet:1", "A1")
	assert.EqualError(t, err, ErrSheetNameInvalid.Error())
	// Non-existent sheet
	_, err = f.GetChart("SheetN", "A1")
	assert.EqualError(t, err, "sheet SheetN does not exist")
	// Invalid cell reference
	_, err = f.GetChart(sheet, "")
	assert.EqualError(t, err, newCellNameToCoordinatesError("", newInvalidCellNameError("")).Error())
	// No drawing on sheet
	f2 := NewFile()
	charts, err = f2.GetChart("Sheet1", "A1")
	assert.NoError(t, err)
	assert.Nil(t, charts)
	assert.NoError(t, f2.Close())

	assert.NoError(t, f.SaveAs(filepath.Join("test", "TestGetChart.xlsx")))
	assert.NoError(t, f.Close())
}

func TestSetChart(t *testing.T) {
	f := NewFile()
	sheet := f.GetSheetName(0)

	for cell, v := range map[string]interface{}{
		"A1": nil, "B1": "Apple", "C1": "Orange", "D1": "Pear",
		"A2": "Small", "B2": 2, "C2": 3, "D2": 3,
		"A3": "Normal", "B3": 5, "C3": 2, "D3": 4,
		"A4": "Large", "B4": 6, "C4": 7, "D4": 8,
	} {
		assert.NoError(t, f.SetCellValue(sheet, cell, v))
	}

	series := []ChartSeries{
		{Name: "Sheet1!$A$2", Categories: "Sheet1!$B$1:$D$1", Values: "Sheet1!$B$2:$D$2"},
		{Name: "Sheet1!$A$3", Categories: "Sheet1!$B$1:$D$1", Values: "Sheet1!$B$3:$D$3"},
		{Name: "Sheet1!$A$4", Categories: "Sheet1!$B$1:$D$1", Values: "Sheet1!$B$4:$D$4"},
	}

	// Add a Col chart
	assert.NoError(t, f.AddChart(sheet, "E1", &Chart{
		Type:   Col,
		Series: series,
		Title:  []RichTextRun{{Text: "Original Chart"}},
	}))
	charts, err := f.GetChart(sheet, "E1")
	assert.NoError(t, err)
	assert.Len(t, charts, 1)
	assert.Equal(t, Col, charts[0].Type)
	assert.Equal(t, "Original Chart", charts[0].Title[0].Text)

	// SetChart: change type to Line and update title
	assert.NoError(t, f.SetChart(sheet, "E1", &Chart{
		Type:   Line,
		Series: series,
		Title:  []RichTextRun{{Text: "Updated Chart"}},
	}))
	charts, err = f.GetChart(sheet, "E1")
	assert.NoError(t, err)
	assert.Len(t, charts, 1)
	assert.Equal(t, Line, charts[0].Type)
	assert.Equal(t, "Updated Chart", charts[0].Title[0].Text)
	assert.Len(t, charts[0].Series, 3)

	// SetChart: update series data
	newSeries := []ChartSeries{
		{Name: "Sheet1!$A$2", Categories: "Sheet1!$B$1:$D$1", Values: "Sheet1!$B$2:$D$2"},
	}
	assert.NoError(t, f.SetChart(sheet, "E1", &Chart{
		Type:   Pie,
		Series: newSeries,
		Title:  []RichTextRun{{Text: "Pie Chart"}},
	}))
	charts, err = f.GetChart(sheet, "E1")
	assert.NoError(t, err)
	assert.Len(t, charts, 1)
	assert.Equal(t, Pie, charts[0].Type)
	assert.Len(t, charts[0].Series, 1)

	// SetChart: replace single chart with combo chart
	assert.NoError(t, f.SetChart(sheet, "E1", &Chart{
		Type:   Col,
		Series: series[:2],
		Title:  []RichTextRun{{Text: "Combo Chart"}},
	}, &Chart{
		Type:   Line,
		Series: series[2:],
	}))
	charts, err = f.GetChart(sheet, "E1")
	assert.NoError(t, err)
	assert.Len(t, charts, 2)
	assert.Equal(t, Col, charts[0].Type)
	assert.Equal(t, Line, charts[1].Type)
	assert.Equal(t, "Combo Chart", charts[0].Title[0].Text)

	// SetChart: on cell with no chart falls back to AddChart
	assert.NoError(t, f.SetChart(sheet, "M1", &Chart{
		Type:   Bar,
		Series: series,
		Title:  []RichTextRun{{Text: "New Bar Chart"}},
	}))
	charts, err = f.GetChart(sheet, "M1")
	assert.NoError(t, err)
	assert.Len(t, charts, 1)
	assert.Equal(t, Bar, charts[0].Type)

	// SetChart: on sheet with no drawings falls back to AddChart
	_, err = f.NewSheet("Sheet2")
	assert.NoError(t, err)
	assert.NoError(t, f.SetChart("Sheet2", "A1", &Chart{
		Type:   Doughnut,
		Series: series,
	}))
	charts, err = f.GetChart("Sheet2", "A1")
	assert.NoError(t, err)
	assert.Len(t, charts, 1)
	assert.Equal(t, Doughnut, charts[0].Type)

	// SetChart: update plot area options
	assert.NoError(t, f.SetChart(sheet, "E1", &Chart{
		Type:   Col,
		Series: series,
		PlotArea: ChartPlotArea{
			ShowVal:     true,
			ShowCatName: true,
		},
		Legend: ChartLegend{Position: "top"},
	}))
	charts, err = f.GetChart(sheet, "E1")
	assert.NoError(t, err)
	assert.True(t, charts[0].PlotArea.ShowVal)
	assert.True(t, charts[0].PlotArea.ShowCatName)
	assert.Equal(t, "top", charts[0].Legend.Position)

	// Save, reopen, verify persistence
	assert.NoError(t, f.SaveAs(filepath.Join("test", "TestSetChart.xlsx")))

	f2, err := OpenFile(filepath.Join("test", "TestSetChart.xlsx"))
	assert.NoError(t, err)
	charts, err = f2.GetChart(sheet, "E1")
	assert.NoError(t, err)
	assert.Len(t, charts, 1)
	assert.Equal(t, Col, charts[0].Type)
	assert.True(t, charts[0].PlotArea.ShowVal)
	assert.NoError(t, f2.Close())

	// Error: invalid sheet name
	assert.EqualError(t, f.SetChart("Sheet:1", "A1", &Chart{Type: Col, Series: series}), ErrSheetNameInvalid.Error())
	// Error: invalid cell reference
	assert.EqualError(t, f.SetChart(sheet, "", &Chart{Type: Col, Series: series}), newCellNameToCoordinatesError("", newInvalidCellNameError("")).Error())
	// Error: non-existent sheet
	assert.EqualError(t, f.SetChart("SheetN", "A1", &Chart{Type: Col, Series: series}), "sheet SheetN does not exist")
	// Error: nil chart options
	assert.EqualError(t, f.SetChart(sheet, "E1", nil), ErrParameterInvalid.Error())

	assert.NoError(t, f.Close())
}

func TestChartAnchorPoint(t *testing.T) {
	f := NewFile()
	sheet := f.GetSheetName(0)

	for cell, v := range map[string]interface{}{
		"A1": nil, "B1": "Apple", "C1": "Orange", "D1": "Pear",
		"A2": "Small", "B2": 2, "C2": 3, "D2": 3,
		"A3": "Normal", "B3": 5, "C3": 2, "D3": 4,
		"A4": "Large", "B4": 6, "C4": 7, "D4": 8,
	} {
		assert.NoError(t, f.SetCellValue(sheet, cell, v))
	}

	series := []ChartSeries{
		{Name: "Sheet1!$A$2", Categories: "Sheet1!$B$1:$D$1", Values: "Sheet1!$B$2:$D$2"},
		{Name: "Sheet1!$A$3", Categories: "Sheet1!$B$1:$D$1", Values: "Sheet1!$B$3:$D$3"},
		{Name: "Sheet1!$A$4", Categories: "Sheet1!$B$1:$D$1", Values: "Sheet1!$B$4:$D$4"},
	}

	fromPt := &AnchorPoint{Col: 4, ColOff: 95250, Row: 0, RowOff: 47625}
	toPt := &AnchorPoint{Col: 12, ColOff: 304800, Row: 15, RowOff: 152400}

	// Test AddChart with explicit From/To anchor points
	assert.NoError(t, f.AddChart(sheet, "E1", &Chart{
		Type:   Col,
		Series: series,
		Title:  []RichTextRun{{Text: "Anchor Test"}},
		Format: GraphicOptions{
			From: fromPt,
			To:   toPt,
		},
	}))

	// Test GetChart returns correct From/To
	charts, err := f.GetChart(sheet, "E1")
	assert.NoError(t, err)
	assert.NotNil(t, charts)
	assert.Len(t, charts, 1)
	assert.Equal(t, Col, charts[0].Type)
	assert.NotNil(t, charts[0].Format.From)
	assert.NotNil(t, charts[0].Format.To)
	assert.Equal(t, 4, charts[0].Format.From.Col)
	assert.Equal(t, 95250, charts[0].Format.From.ColOff)
	assert.Equal(t, 0, charts[0].Format.From.Row)
	assert.Equal(t, 47625, charts[0].Format.From.RowOff)
	assert.Equal(t, 12, charts[0].Format.To.Col)
	assert.Equal(t, 304800, charts[0].Format.To.ColOff)
	assert.Equal(t, 15, charts[0].Format.To.Row)
	assert.Equal(t, 152400, charts[0].Format.To.RowOff)

	// Test persistence: save and reopen
	assert.NoError(t, f.SaveAs(filepath.Join("test", "TestChartAnchorPoint.xlsx")))

	f2, err := OpenFile(filepath.Join("test", "TestChartAnchorPoint.xlsx"))
	assert.NoError(t, err)
	charts, err = f2.GetChart(sheet, "E1")
	assert.NoError(t, err)
	assert.NotNil(t, charts)
	assert.Len(t, charts, 1)
	assert.NotNil(t, charts[0].Format.From)
	assert.NotNil(t, charts[0].Format.To)
	assert.Equal(t, 4, charts[0].Format.From.Col)
	assert.Equal(t, 95250, charts[0].Format.From.ColOff)
	assert.Equal(t, 0, charts[0].Format.From.Row)
	assert.Equal(t, 47625, charts[0].Format.From.RowOff)
	assert.Equal(t, 12, charts[0].Format.To.Col)
	assert.Equal(t, 304800, charts[0].Format.To.ColOff)
	assert.Equal(t, 15, charts[0].Format.To.Row)
	assert.Equal(t, 152400, charts[0].Format.To.RowOff)
	assert.NoError(t, f2.Close())

	// Test SetChart updates From/To anchor position (keep From.Col/Row same for lookup)
	newFrom := &AnchorPoint{Col: 4, ColOff: 19050, Row: 0, RowOff: 9525}
	newTo := &AnchorPoint{Col: 14, ColOff: 190500, Row: 17, RowOff: 95250}
	assert.NoError(t, f.SetChart(sheet, "E1", &Chart{
		Type:   Col,
		Series: series,
		Title:  []RichTextRun{{Text: "Moved Chart"}},
		Format: GraphicOptions{
			From: newFrom,
			To:   newTo,
		},
	}))
	charts, err = f.GetChart(sheet, "E1")
	assert.NoError(t, err)
	assert.NotNil(t, charts)
	assert.Len(t, charts, 1)
	assert.NotNil(t, charts[0].Format.From)
	assert.NotNil(t, charts[0].Format.To)
	assert.Equal(t, 4, charts[0].Format.From.Col)
	assert.Equal(t, 19050, charts[0].Format.From.ColOff)
	assert.Equal(t, 0, charts[0].Format.From.Row)
	assert.Equal(t, 9525, charts[0].Format.From.RowOff)
	assert.Equal(t, 14, charts[0].Format.To.Col)
	assert.Equal(t, 190500, charts[0].Format.To.ColOff)
	assert.Equal(t, 17, charts[0].Format.To.Row)
	assert.Equal(t, 95250, charts[0].Format.To.RowOff)

	// Test SetChart with From/To after save/reopen (file-loaded anchor)
	assert.NoError(t, f.SaveAs(filepath.Join("test", "TestChartAnchorPoint.xlsx")))
	f4, err := OpenFile(filepath.Join("test", "TestChartAnchorPoint.xlsx"))
	assert.NoError(t, err)
	reopenFrom := &AnchorPoint{Col: 4, ColOff: 10000, Row: 0, RowOff: 20000}
	reopenTo := &AnchorPoint{Col: 11, ColOff: 30000, Row: 14, RowOff: 40000}
	assert.NoError(t, f4.SetChart(sheet, "E1", &Chart{
		Type:   Col,
		Series: series,
		Title:  []RichTextRun{{Text: "Reopened Moved Chart"}},
		Format: GraphicOptions{
			From: reopenFrom,
			To:   reopenTo,
		},
	}))
	charts, err = f4.GetChart(sheet, "E1")
	assert.NoError(t, err)
	assert.NotNil(t, charts)
	assert.Len(t, charts, 1)
	assert.NotNil(t, charts[0].Format.From)
	assert.NotNil(t, charts[0].Format.To)
	assert.Equal(t, 4, charts[0].Format.From.Col)
	assert.Equal(t, 10000, charts[0].Format.From.ColOff)
	assert.Equal(t, 0, charts[0].Format.From.Row)
	assert.Equal(t, 20000, charts[0].Format.From.RowOff)
	assert.Equal(t, 11, charts[0].Format.To.Col)
	assert.Equal(t, 30000, charts[0].Format.To.ColOff)
	assert.Equal(t, 14, charts[0].Format.To.Row)
	assert.Equal(t, 40000, charts[0].Format.To.RowOff)
	assert.NoError(t, f4.Close())

	// Test AddChart without From/To still works (backward compatibility)
	f3 := NewFile()
	for cell, v := range map[string]interface{}{
		"A1": nil, "B1": "Apple", "C1": "Orange", "D1": "Pear",
		"A2": "Small", "B2": 2, "C2": 3, "D2": 3,
	} {
		assert.NoError(t, f3.SetCellValue(f3.GetSheetName(0), cell, v))
	}
	assert.NoError(t, f3.AddChart(f3.GetSheetName(0), "E1", &Chart{
		Type:   Col,
		Series: series,
		Title:  []RichTextRun{{Text: "No Anchor"}},
	}))
	charts, err = f3.GetChart(f3.GetSheetName(0), "E1")
	assert.NoError(t, err)
	assert.NotNil(t, charts)
	assert.Len(t, charts, 1)
	// From/To should still be populated from auto-calculated positions
	assert.NotNil(t, charts[0].Format.From)
	assert.NotNil(t, charts[0].Format.To)
	assert.NoError(t, f3.Close())

	// Test RID is populated by GetChart
	assert.NotEmpty(t, fromPt) // reuse fromPt from above
	charts, err = f.GetChart(sheet, "E1")
	assert.NoError(t, err)
	assert.Len(t, charts, 1)
	assert.NotEmpty(t, charts[0].RID)

	// Test SetChart with RID-based lookup
	rid := charts[0].RID
	assert.NoError(t, f.SetChart(sheet, "E1", &Chart{
		Type:   Bar,
		Series: series,
		Title:  []RichTextRun{{Text: "Updated via RID"}},
		RID:    rid,
		Format: GraphicOptions{
			From: &AnchorPoint{Col: 4, ColOff: 11111, Row: 0, RowOff: 22222},
			To:   &AnchorPoint{Col: 10, ColOff: 33333, Row: 12, RowOff: 44444},
		},
	}))
	charts, err = f.GetChart(sheet, "E1")
	assert.NoError(t, err)
	assert.Len(t, charts, 1)
	assert.Equal(t, Bar, charts[0].Type)
	assert.Equal(t, 11111, charts[0].Format.From.ColOff)
	assert.Equal(t, 22222, charts[0].Format.From.RowOff)
	assert.Equal(t, 33333, charts[0].Format.To.ColOff)
	assert.Equal(t, 44444, charts[0].Format.To.RowOff)
	assert.Equal(t, rid, charts[0].RID)

	assert.NoError(t, f.Close())
}

func TestChartMultipleOnSameCell(t *testing.T) {
	f := NewFile()
	sheet := f.GetSheetName(0)

	for cell, v := range map[string]interface{}{
		"A1": nil, "B1": "Apple", "C1": "Orange", "D1": "Pear",
		"A2": "Small", "B2": 2, "C2": 3, "D2": 3,
		"A3": "Normal", "B3": 5, "C3": 2, "D3": 4,
		"A4": "Large", "B4": 6, "C4": 7, "D4": 8,
	} {
		assert.NoError(t, f.SetCellValue(sheet, cell, v))
	}

	series := []ChartSeries{
		{Name: "Sheet1!$A$2", Categories: "Sheet1!$B$1:$D$1", Values: "Sheet1!$B$2:$D$2"},
		{Name: "Sheet1!$A$3", Categories: "Sheet1!$B$1:$D$1", Values: "Sheet1!$B$3:$D$3"},
		{Name: "Sheet1!$A$4", Categories: "Sheet1!$B$1:$D$1", Values: "Sheet1!$B$4:$D$4"},
	}

	// Add two charts to the same cell "E1"
	assert.NoError(t, f.AddChart(sheet, "E1", &Chart{
		Type:   Col,
		Series: series,
		Title:  []RichTextRun{{Text: "First Chart"}},
	}))
	assert.NoError(t, f.AddChart(sheet, "E1", &Chart{
		Type:   Line,
		Series: series,
		Title:  []RichTextRun{{Text: "Second Chart"}},
	}))

	// GetChart should return both charts at E1 (each in its own []*Chart)
	charts, err := f.GetChart(sheet, "E1")
	assert.NoError(t, err)
	assert.NotNil(t, charts)
	assert.Len(t, charts, 2)
	// Charts should have different RIDs
	assert.NotEmpty(t, charts[0].RID)
	assert.NotEmpty(t, charts[1].RID)
	assert.NotEqual(t, charts[0].RID, charts[1].RID)
	// Check types
	assert.Equal(t, Col, charts[0].Type)
	assert.Equal(t, Line, charts[1].Type)

	// Test GetCharts returns all charts on the sheet
	allCharts, err := f.GetCharts(sheet)
	assert.NoError(t, err)
	assert.Len(t, allCharts, 2)

	// Test GetChartByRID returns the correct chart
	rid1 := charts[0].RID
	rid2 := charts[1].RID
	chart1, err := f.GetChartByRID(sheet, rid1)
	assert.NoError(t, err)
	assert.NotNil(t, chart1)
	assert.Equal(t, Col, chart1.Type)
	assert.Equal(t, rid1, chart1.RID)
	chart2, err := f.GetChartByRID(sheet, rid2)
	assert.NoError(t, err)
	assert.NotNil(t, chart2)
	assert.Equal(t, Line, chart2.Type)
	assert.Equal(t, rid2, chart2.RID)

	// Test GetChartByRID with non-existent RID
	chart, err := f.GetChartByRID(sheet, "rIdNonExistent")
	assert.NoError(t, err)
	assert.Nil(t, chart)

	// Test SetChart by RID: update only the second chart
	assert.NoError(t, f.SetChart(sheet, "E1", &Chart{
		Type:   Bar,
		Series: series,
		Title:  []RichTextRun{{Text: "Updated Second Chart"}},
		RID:    rid2,
	}))
	// Verify: first chart unchanged, second chart updated
	chart1, err = f.GetChartByRID(sheet, rid1)
	assert.NoError(t, err)
	assert.Equal(t, Col, chart1.Type)
	assert.Equal(t, "First Chart", chart1.Title[0].Text)
	chart2, err = f.GetChartByRID(sheet, rid2)
	assert.NoError(t, err)
	assert.Equal(t, Bar, chart2.Type)
	assert.Equal(t, "Updated Second Chart", chart2.Title[0].Text)

	// Test DeleteChartByRID: delete the first chart
	assert.NoError(t, f.DeleteChartByRID(sheet, rid1))
	// First chart should be gone
	chart1, err = f.GetChartByRID(sheet, rid1)
	assert.NoError(t, err)
	assert.Nil(t, chart1)
	// Second chart should still exist
	chart2, err = f.GetChartByRID(sheet, rid2)
	assert.NoError(t, err)
	assert.NotNil(t, chart2)
	assert.Equal(t, Bar, chart2.Type)
	// GetCharts should now show only 1
	allCharts, err = f.GetCharts(sheet)
	assert.NoError(t, err)
	assert.Len(t, allCharts, 1)
	// GetChart at E1 should show only 1
	charts, err = f.GetChart(sheet, "E1")
	assert.NoError(t, err)
	assert.Len(t, charts, 1)

	// Test DeleteChartByRID: delete the remaining chart
	assert.NoError(t, f.DeleteChartByRID(sheet, rid2))
	allCharts, err = f.GetCharts(sheet)
	assert.NoError(t, err)
	assert.Nil(t, allCharts)
	charts, err = f.GetChart(sheet, "E1")
	assert.NoError(t, err)
	assert.Nil(t, charts)

	// Test DeleteChartByRID with non-existent RID (no error)
	assert.NoError(t, f.DeleteChartByRID(sheet, "rIdNonExistent"))

	// Test error cases
	// Invalid sheet name
	assert.EqualError(t, f.DeleteChartByRID("Sheet:1", "rId1"), ErrSheetNameInvalid.Error())
	_, err = f.GetCharts("Sheet:1")
	assert.EqualError(t, err, ErrSheetNameInvalid.Error())
	_, err = f.GetChartByRID("Sheet:1", "rId1")
	assert.EqualError(t, err, ErrSheetNameInvalid.Error())
	// Non-existent sheet
	assert.EqualError(t, f.DeleteChartByRID("SheetN", "rId1"), "sheet SheetN does not exist")
	_, err = f.GetCharts("SheetN")
	assert.EqualError(t, err, "sheet SheetN does not exist")
	_, err = f.GetChartByRID("SheetN", "rId1")
	assert.EqualError(t, err, "sheet SheetN does not exist")
	// No drawing on sheet
	f2 := NewFile()
	assert.NoError(t, f2.DeleteChartByRID("Sheet1", "rId1"))
	charts2, err := f2.GetCharts("Sheet1")
	assert.NoError(t, err)
	assert.Nil(t, charts2)
	chart, err = f2.GetChartByRID("Sheet1", "rId1")
	assert.NoError(t, err)
	assert.Nil(t, chart)
	assert.NoError(t, f2.Close())

	assert.NoError(t, f.SaveAs(filepath.Join("test", "TestChartMultipleOnSameCell.xlsx")))
	assert.NoError(t, f.Close())
}

func TestGetCharts(t *testing.T) {
	f := NewFile()
	sheet := f.GetSheetName(0)

	for cell, v := range map[string]interface{}{
		"A1": nil, "B1": "Apple", "C1": "Orange", "D1": "Pear",
		"A2": "Small", "B2": 2, "C2": 3, "D2": 3,
		"A3": "Normal", "B3": 5, "C3": 2, "D3": 4,
		"A4": "Large", "B4": 6, "C4": 7, "D4": 8,
	} {
		assert.NoError(t, f.SetCellValue(sheet, cell, v))
	}

	series := []ChartSeries{
		{Name: "Sheet1!$A$2", Categories: "Sheet1!$B$1:$D$1", Values: "Sheet1!$B$2:$D$2"},
		{Name: "Sheet1!$A$3", Categories: "Sheet1!$B$1:$D$1", Values: "Sheet1!$B$3:$D$3"},
	}

	// No charts initially
	charts, err := f.GetCharts(sheet)
	assert.NoError(t, err)
	assert.Nil(t, charts)

	// Add charts at different cells
	assert.NoError(t, f.AddChart(sheet, "E1", &Chart{
		Type:   Col,
		Series: series,
		Title:  []RichTextRun{{Text: "Chart at E1"}},
	}))
	assert.NoError(t, f.AddChart(sheet, "M1", &Chart{
		Type:   Line,
		Series: series,
		Title:  []RichTextRun{{Text: "Chart at M1"}},
	}))
	assert.NoError(t, f.AddChart(sheet, "E16", &Chart{
		Type:   Pie,
		Series: series,
		Title:  []RichTextRun{{Text: "Chart at E16"}},
	}))

	// GetCharts returns all 3 charts
	charts, err = f.GetCharts(sheet)
	assert.NoError(t, err)
	assert.Len(t, charts, 3)
	// All should have non-empty RIDs
	for _, c := range charts {
		assert.NotEmpty(t, c.RID)
	}

	// Verify persistence after save/reopen
	assert.NoError(t, f.SaveAs(filepath.Join("test", "TestGetCharts.xlsx")))
	f2, err := OpenFile(filepath.Join("test", "TestGetCharts.xlsx"))
	assert.NoError(t, err)
	charts, err = f2.GetCharts(sheet)
	assert.NoError(t, err)
	assert.Len(t, charts, 3)
	assert.NoError(t, f2.Close())

	assert.NoError(t, f.Close())
}

func TestGetChartsAxisTitles(t *testing.T) {
	f := NewFile()
	sheet := f.GetSheetName(0)

	for cell, v := range map[string]interface{}{
		"A1": nil, "B1": "Sales",
		"A2": "Q1", "B2": 10,
		"A3": "Q2", "B3": 20,
	} {
		assert.NoError(t, f.SetCellValue(sheet, cell, v))
	}

	assert.NoError(t, f.AddChart(sheet, "E1", &Chart{
		Type: Col,
		Series: []ChartSeries{
			{Name: "Sheet1!$B$1", Categories: "Sheet1!$A$2:$A$3", Values: "Sheet1!$B$2:$B$3"},
		},
		Title: []RichTextRun{{Text: "Quarterly Sales"}},
		XAxis: ChartAxis{Title: []RichTextRun{{Text: "Quarter"}}},
		YAxis: ChartAxis{Title: []RichTextRun{{Text: "Amount"}}},
	}))

	path := filepath.Join("test", "TestGetChartsAxisTitles.xlsx")
	assert.NoError(t, f.SaveAs(path))

	reopened, err := OpenFile(path)
	assert.NoError(t, err)
	if err != nil {
		assert.NoError(t, f.Close())
		return
	}

	charts, err := reopened.GetCharts(sheet)
	assert.NoError(t, err)
	if !assert.Len(t, charts, 1) {
		_ = reopened.Close()
		assert.NoError(t, f.Close())
		return
	}
	if assert.Len(t, charts[0].XAxis.Title, 1) {
		assert.Equal(t, "Quarter", charts[0].XAxis.Title[0].Text)
	}
	if assert.Len(t, charts[0].YAxis.Title, 1) {
		assert.Equal(t, "Amount", charts[0].YAxis.Title[0].Text)
	}

	assert.NoError(t, reopened.Close())
	assert.NoError(t, f.Close())
}
