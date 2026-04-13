package excelize

import (
	"strings"
	"testing"
)

func TestNestedFunctionArgumentIsolation(t *testing.T) {
	f := NewFile()

	productImages := `[{"image_type": "product_image_url", "url": "https://statics.maybe.ai/videos/chunse_weiyi_nv.jpg"}, {"image_type": "product_attribute_url", "url": "https://statics.maybe.ai/videos/weiyi_nv_attributes.png"}]`
	if err := f.SetCellValue("Sheet1", "B2", productImages); err != nil {
		t.Fatal(err)
	}

	formula := `"[INPUT DATA STREAM] * **[PRODUCT IMAGES] (The Subject):**"&IF((LEN(B2)-LEN(SUBSTITUTE(B2,"https","")))/5<=1,"image1","image1, image"&TEXT((LEN(B2)-LEN(SUBSTITUTE(B2,"https","")))/5,"0"))&"PRODUCT URLS: "&B2&"**"`
	if err := f.SetCellFormula("Sheet1", "L2", formula); err != nil {
		t.Fatal(err)
	}

	got, err := f.CalcCellValue("Sheet1", "L2", Options{RawCellValue: true})
	if err != nil {
		t.Fatalf("CalcCellValue failed: %v", err)
	}
	if !strings.Contains(got, "image1, image2") {
		t.Fatalf("CalcCellValue returned unexpected result: %q", got)
	}

	if err := f.RecalculateAllWithDependency(); err != nil {
		t.Fatalf("RecalculateAllWithDependency failed: %v", err)
	}

	recalcValue, err := f.GetCellValue("Sheet1", "L2", Options{RawCellValue: true})
	if err != nil {
		t.Fatalf("GetCellValue after recalc failed: %v", err)
	}
	if got != recalcValue {
		t.Fatalf("expected recalc result %q, got %q", got, recalcValue)
	}
}
