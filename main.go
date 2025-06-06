package main

import (
	"time"

	excel2structpoc "github.com/kittanutp/playground/excel-2-struct-poc"
)

type Data struct {
	Term           int       `json:"term"`
	SubAccount     string    `json:"sub_account"`
	CompanyName    string    `json:"company_name"`
	IsAffiliation  bool      `json:"is_affiliation"`
	PaymentDueDate time.Time `json:"payment_due_date"`
}

func main() {
	excel2structpoc.Run()
}
