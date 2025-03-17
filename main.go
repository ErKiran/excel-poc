package main

import (
	"encoding/json"
	"fmt"
	"log"
	"net/http"
	"os"
	"path/filepath"

	"github.com/xuri/excelize/v2"
)

const (
	stockAPIURL = "https://www.onlinekhabar.com/smtm/home/trending"
	dataDir     = "data"
	excelFile   = "stock_data.xlsx"
	sheetName   = "Stock Data"
)

type StockAPIResponse struct {
	Response []StockData `json:"response"`
}

type StockData struct {
	Ticker           string  `json:"ticker"`
	TickerName       string  `json:"ticker_name"`
	LatestPrice      string  `json:"latest_price"`
	PointsChange     float64 `json:"points_change"`
	PercentageChange float64 `json:"percentage_change"`
	TradedOfMktCap   float64 `json:"traded_of_mkt_cap"`
}

type StockService struct {
	apiURL string
}

func NewStockService(apiURL string) *StockService {
	return &StockService{apiURL: apiURL}
}

func (s *StockService) FetchStockData() ([]StockData, error) {
	resp, err := http.Get(s.apiURL)
	if err != nil {
		return nil, fmt.Errorf("failed to fetch stock data: %w", err)
	}
	defer resp.Body.Close()

	if resp.StatusCode != http.StatusOK {
		return nil, fmt.Errorf("API request failed with status code: %d", resp.StatusCode)
	}

	var stockResponse StockAPIResponse
	if err := json.NewDecoder(resp.Body).Decode(&stockResponse); err != nil {
		return nil, fmt.Errorf("failed to decode JSON response: %w", err)
	}

	return stockResponse.Response, nil
}

type ExcelGenerator struct {
	filePath string
}

func NewExcelGenerator(directory, filename string) *ExcelGenerator {
	return &ExcelGenerator{
		filePath: filepath.Join(directory, filename),
	}
}

func (e *ExcelGenerator) GenerateExcelFile(stocks []StockData) error {
	f := excelize.NewFile()
	f.SetSheetName("Sheet1", sheetName)

	headers := []string{"Ticker", "Ticker Name", "Latest Price", "Points Change", "Percentage Change", "Traded Of Mkt Cap"}
	for col, header := range headers {
		cell := fmt.Sprintf("%c1", 'A'+col)
		f.SetCellValue(sheetName, cell, header)
	}

	for i, stock := range stocks {
		f.SetCellValue(sheetName, fmt.Sprintf("A%d", i+2), stock.Ticker)
		f.SetCellValue(sheetName, fmt.Sprintf("B%d", i+2), stock.TickerName)
		f.SetCellValue(sheetName, fmt.Sprintf("C%d", i+2), stock.LatestPrice)
		f.SetCellValue(sheetName, fmt.Sprintf("D%d", i+2), stock.PointsChange)
		f.SetCellValue(sheetName, fmt.Sprintf("E%d", i+2), stock.PercentageChange)
		f.SetCellValue(sheetName, fmt.Sprintf("F%d", i+2), stock.TradedOfMktCap)
	}

	if err := os.MkdirAll(filepath.Dir(e.filePath), os.ModePerm); err != nil {
		return fmt.Errorf("failed to create directory: %w", err)
	}

	if err := f.SaveAs(e.filePath); err != nil {
		return fmt.Errorf("failed to save Excel file: %w", err)
	}

	fmt.Println("Excel file saved at:", e.filePath)
	return nil
}

func main() {

	stockService := NewStockService(stockAPIURL)
	excelGenerator := NewExcelGenerator(dataDir, excelFile)

	stocks, err := stockService.FetchStockData()
	if err != nil {
		log.Fatalf("Error fetching stock data: %v", err)
	}

	if err := excelGenerator.GenerateExcelFile(stocks); err != nil {
		log.Fatalf("Error generating Excel file: %v", err)
	}
}
