package excelize

import (
	"strconv"
	"sync"
)

// WorksheetCache 统一的工作表缓存，按 sheet 组织
// 用于存储所有单元格的值（包括原始值和计算结果）
// Phase 1 重构：改为存储 formulaArg 以保留类型信息
type WorksheetCache struct {
	mu    sync.RWMutex
	cache map[string]map[string]formulaArg // map[sheetName]map[cellRef]formulaArg
}

// NewWorksheetCache 创建新的工作表缓存
func NewWorksheetCache() *WorksheetCache {
	return &WorksheetCache{
		cache: make(map[string]map[string]formulaArg),
	}
}

// Get 获取单元格的值
// 返回 formulaArg 和是否存在的标志
func (wc *WorksheetCache) Get(sheet, cell string) (formulaArg, bool) {
	wc.mu.RLock()
	defer wc.mu.RUnlock()

	if sheetCache, ok := wc.cache[sheet]; ok {
		value, exists := sheetCache[cell]
		return value, exists
	}
	return newEmptyFormulaArg(), false
}

// Set 设置单元格的值
func (wc *WorksheetCache) Set(sheet, cell string, value formulaArg) {
	wc.mu.Lock()
	defer wc.mu.Unlock()

	if _, ok := wc.cache[sheet]; !ok {
		wc.cache[sheet] = make(map[string]formulaArg)
	}
	wc.cache[sheet][cell] = value
}

// Delete removes a single cached cell value.
func (wc *WorksheetCache) Delete(sheet, cell string) {
	wc.mu.Lock()
	defer wc.mu.Unlock()

	if sheetCache, ok := wc.cache[sheet]; ok {
		delete(sheetCache, cell)
		if len(sheetCache) == 0 {
			delete(wc.cache, sheet)
		}
	}
}

// GetSheet 获取整个 sheet 的数据（用于批量操作）
// 返回 map[cellRef]formulaArg
func (wc *WorksheetCache) GetSheet(sheet string) map[string]formulaArg {
	wc.mu.RLock()
	defer wc.mu.RUnlock()

	if sheetCache, ok := wc.cache[sheet]; ok {
		// 返回副本，避免并发修改
		result := make(map[string]formulaArg, len(sheetCache))
		for k, v := range sheetCache {
			result[k] = v
		}
		return result
	}
	return make(map[string]formulaArg)
}

// GetCacheStats 返回缓存统计信息（用于调试）
func (wc *WorksheetCache) GetCacheStats() map[string]int {
	wc.mu.RLock()
	defer wc.mu.RUnlock()

	stats := make(map[string]int)
	total := 0
	for sheet, sheetCache := range wc.cache {
		count := len(sheetCache)
		stats[sheet] = count
		total += count
	}
	stats["_total"] = total
	return stats
}

// inferCellValueType 根据单元格的原始值推断其类型并转换为 formulaArg
// 这是 Phase 1 的核心：保留类型信息而不是只存字符串
func inferCellValueType(val string, cellType CellType) formulaArg {
	// 空字符串：保持为字符串类型（在比较时会被特殊处理为 0）
	if val == "" {
		return newStringFormulaArg("")
	}

	// 根据单元格类型判断
	switch cellType {
	case CellTypeBool:
		// 布尔类型
		return newBoolFormulaArg(val == "1" || val == "TRUE" || val == "true")

	case CellTypeNumber, CellTypeUnset:
		// 数值类型：尝试解析为数字
		if num, err := strconv.ParseFloat(val, 64); err == nil {
			return newNumberFormulaArg(num)
		}
		// 解析失败，作为字符串
		return newStringFormulaArg(val)

	default:
		// 其他类型（字符串、日期等）：保持原样为字符串
		// 重要：不要尝试解析为数字，因为 Excel 中字符串 "0" 和数字 0 是不同的类型
		return newStringFormulaArg(val)
	}
}

// LoadSheet 加载整个 sheet 的数据到缓存
// Phase 1 改进：读取时立即转换为 formulaArg，保留类型信息
func (wc *WorksheetCache) LoadSheet(f *File, sheet string) error {
	// 先确保 map 初始化
	wc.mu.Lock()
	if _, ok := wc.cache[sheet]; !ok {
		wc.cache[sheet] = make(map[string]formulaArg)
	}
	wc.mu.Unlock()

	ws, err := f.workSheetReader(sheet)
	if err != nil || ws == nil || ws.SheetData.Row == nil {
		return err
	}

	for _, row := range ws.SheetData.Row {
		for _, cell := range row.C {
			if cell.F != nil {
				// 公式单元格的值通过计算阶段缓存
				continue
			}
			// 读取原始值（不格式化）
			val, err := f.GetCellValue(sheet, cell.R, Options{RawCellValue: true})
			if err != nil || val == "" {
				// 空值也需要缓存（作为空字符串）
				if val == "" && err == nil {
					arg := newStringFormulaArg("")
					wc.Set(sheet, cell.R, arg)
				}
				continue
			}

			// 获取单元格类型
			cellType, _ := f.GetCellType(sheet, cell.R)

			// 推断类型并转换为 formulaArg
			arg := inferCellValueType(val, cellType)
			wc.Set(sheet, cell.R, arg)
		}
	}
	return nil
}

// Clear 清空缓存
func (wc *WorksheetCache) Clear() {
	wc.mu.Lock()
	defer wc.mu.Unlock()
	wc.cache = make(map[string]map[string]formulaArg)
}

// ClearSheet 清空指定 sheet 的缓存
func (wc *WorksheetCache) ClearSheet(sheet string) {
	wc.mu.Lock()
	defer wc.mu.Unlock()
	delete(wc.cache, sheet)
}

// Len 返回总的缓存单元格数量
func (wc *WorksheetCache) Len() int {
	wc.mu.RLock()
	defer wc.mu.RUnlock()

	total := 0
	for _, sheetCache := range wc.cache {
		total += len(sheetCache)
	}
	return total
}

// SheetLen 返回指定 sheet 的缓存单元格数量
func (wc *WorksheetCache) SheetLen(sheet string) int {
	wc.mu.RLock()
	defer wc.mu.RUnlock()

	if sheetCache, ok := wc.cache[sheet]; ok {
		return len(sheetCache)
	}
	return 0
}
