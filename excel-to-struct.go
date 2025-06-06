package excel2structpoc

import (
	"encoding/json"
	"fmt"
	"strconv"
	"strings"
	"time"

	"github.com/xuri/excelize/v2"
)

const ExcelDateFormat string = "02-01-2006"

type ExcelRow struct {
	SaleChannel  string  `json:"sale_channel"`
	SaleSource   string  `json:"sale_source"`
	SkuCode      string  `json:"sku_code"`
	IsUnlimited  string  `json:"is_unlimited"`
	BufferStock  string  `json:"buffer_stock"`
	BadgeIds     string  `json:"badge_ids"`
	IsActive     string  `json:"is_active"`
	OffStartDate *string `json:"off_start_date"`
	OffEndate    *string `json:"off_end_date"`
}

func (in *ExcelRow) Compose() (*Excel, error) {
	var err error
	// initiate default value here
	result := Excel{
		SaleChannel: in.SaleChannel,
		SaleSource:  in.SaleSource,
		SkuCode:     in.SkuCode,
	}

	// Convert bool
	if in.IsUnlimited != "" {
		result.IsUnlimited, err = strconv.ParseBool(in.IsUnlimited)
		if err != nil {
			return nil, fmt.Errorf("invalid is_unlimited '%s': %+v", in.IsUnlimited, err)
		}
	}

	if in.IsActive != "" {
		result.IsActive, err = strconv.ParseBool(in.IsActive)
		if err != nil {
			return nil, fmt.Errorf("invalid is_active '%s': %+v", in.IsActive, err)
		}
	}

	// Convert BufferStock
	if in.BufferStock != "" {
		result.BufferStock, err = strconv.Atoi(in.BufferStock)
		if err != nil {
			return nil, fmt.Errorf("invalid buffer_stock '%s': %+v", in.BufferStock, err)
		}
	}

	// Parse BadgeIds
	if in.BadgeIds != "" {
		result.BadgeIds = strings.Split(in.BadgeIds, ",")
		for i := range result.BadgeIds {
			result.BadgeIds[i] = strings.TrimSpace(result.BadgeIds[i])
		}
	}

	// Parse OffStartDate
	if in.OffStartDate != nil && *in.OffStartDate != "" {
		t, err := time.Parse(ExcelDateFormat, *in.OffStartDate)
		if err != nil {
			return nil, fmt.Errorf("invalid off_start_date '%s': %w", *in.OffStartDate, err)
		}
		result.InactiveSchedule.StartDate = &t
	}

	// Parse OffEndate
	if in.OffEndate != nil && *in.OffEndate != "" {
		t, err := time.Parse(ExcelDateFormat, *in.OffEndate)
		if err != nil {
			return nil, fmt.Errorf("invalid off_end_date '%s': %w", *in.OffEndate, err)
		}
		result.InactiveSchedule.EndDate = &t
	}

	return &result, nil
}

type Excel struct {
	SaleChannel      string
	SaleSource       string
	SkuCode          string
	IsUnlimited      bool
	BufferStock      int
	BadgeIds         []string
	IsActive         bool
	InactiveSchedule InactiveSchedule
}

type InactiveSchedule struct {
	StartDate *time.Time
	EndDate   *time.Time
}

func Run() {

	f, err := excelize.OpenFile("excel-2-struct-poc/poc.xlsx")
	if err != nil {
		fmt.Println(err)
		return
	}

	rows, err := f.GetRows("Sheet1")
	if err != nil {
		fmt.Println(err)
		return
	}

	var (
		header = rows[0]
		data   = make([]Excel, 0, len(rows)-1)
	)

	for _, row := range rows[1:] {
		item := make(map[string]any, len(row))
		for idx, v := range row {
			key := header[idx]
			item[key] = v
		}
		o, err := JsonToStruct[ExcelRow](item)
		if err != nil {
			fmt.Println(err)
			return
		}

		res, err := o.Compose()
		if err != nil {
			fmt.Println(err)
			return
		}

		if res != nil {
			data = append(data, *res)
		}
	}

	fmt.Print(data)
}

func JsonToStruct[T any](in map[string]any) (*T, error) {
	var o T
	b, err := json.Marshal(in)
	if err != nil {
		return nil, err
	}
	if err := json.Unmarshal(b, &o); err != nil {
		return nil, err
	}
	return &o, nil
}
