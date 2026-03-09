package main

import (
	"fmt"
	"log"
	"strconv"

	"github.com/Wenchy/requests"
	"github.com/xuri/excelize/v2"
)

type Player struct {
	Ranking  int     `json:"ranking"`
	Name     string  `json:"name"`
	UID      string  `json:"uid"`
	CertName string  `json:"certName"`
	Score    float64 `json:"score"`
}

type Reply struct {
	Code    int  `json:"code"`
	Success bool `json:"success"`
	Data    struct {
		PageIndex      int `json:"pageIndex"`
		PageSize       int `json:"pageSize"`
		TotalPageCount int `json:"totalPageCount"`
		TotalCount     int `json:"totalCount"`
		List           []struct {
			Ranking     int     `json:"ranking"`
			Name        string  `json:"name"`
			AvatarURL   string  `json:"avatarUrl"`
			UID         string  `json:"uid"`
			AreaName    string  `json:"areaName"`
			CertName    string  `json:"certName"`
			Score       float64 `json:"score"`
			PublicScore float64 `json:"publicScore"`
			PublicLevel int     `json:"publicLevel"`
			PublicTime  string  `json:"publicTime"`
		} `json:"list"`
	} `json:"data"`
	Msg string `json:"msg"`
}

// Product defines a mahjong game type with its productId and sheet name.
type Product struct {
	ProductID string
	SheetName string
}

func fetchPage(productId string, pageIndex int) (*Reply, []Player, error) {
	var reply Reply
	_, err := requests.Get(
		"https://wxapi.mahjonget.com/cert/open/score/page",
		requests.ParamPairs(
			"productId", productId,
			"pageIndex", strconv.Itoa(pageIndex),
			"pageSize", "100",
		),
		requests.ToJSON(&reply),
	)
	if err != nil {
		return nil, nil, err
	}

	if !reply.Success {
		return nil, nil, fmt.Errorf("API error: %s", reply.Msg)
	}

	var players []Player
	for _, item := range reply.Data.List {
		players = append(players, Player{
			Ranking:  item.Ranking,
			Name:     item.Name,
			UID:      item.UID,
			CertName: item.CertName,
			Score:    item.Score,
		})
	}
	return &reply, players, nil
}

// fetchAllPlayers fetches all pages for a given productId.
func fetchAllPlayers(productId string) ([]Player, error) {
	firstReply, firstPagePlayers, err := fetchPage(productId, 1)
	if err != nil {
		return nil, fmt.Errorf("failed to fetch first page: %v", err)
	}

	totalPages := firstReply.Data.TotalPageCount
	if totalPages <= 0 {
		totalPages = 1
	}

	var allPlayers []Player
	allPlayers = append(allPlayers, firstPagePlayers...)

	log.Printf("Total pages: %d", totalPages)
	log.Printf("Fetched page 1, players: %d", len(firstPagePlayers))

	for page := 2; page <= totalPages; page++ {
		_, players, err := fetchPage(productId, page)
		if err != nil {
			log.Printf("Warning: Failed to fetch page %d: %v", page, err)
			continue
		}
		allPlayers = append(allPlayers, players...)
		log.Printf("Fetched page %d, players: %d", page, len(players))
	}

	return allPlayers, nil
}

// writeSheet writes player data to a specific sheet in the Excel file.
func writeSheet(f *excelize.File, sheetName string, players []Player) error {
	_, err := f.NewSheet(sheetName)
	if err != nil {
		return err
	}

	// Set headers
	headers := []string{"Ranking", "Name", "UID", "CertName", "Score"}
	for col, header := range headers {
		cell, _ := excelize.CoordinatesToCellName(col+1, 1)
		f.SetCellValue(sheetName, cell, header)
	}

	// Fill data
	for row, player := range players {
		f.SetCellValue(sheetName, fmt.Sprintf("A%d", row+2), player.Ranking)
		f.SetCellValue(sheetName, fmt.Sprintf("B%d", row+2), player.Name)
		f.SetCellValue(sheetName, fmt.Sprintf("C%d", row+2), player.UID)
		f.SetCellValue(sheetName, fmt.Sprintf("D%d", row+2), player.CertName)
		f.SetCellValue(sheetName, fmt.Sprintf("E%d", row+2), player.Score)
	}

	return nil
}

func main() {
	products := []Product{
		{ProductID: "1739485995675418624", SheetName: "国标麻将(MCR)"},
		{ProductID: "1739484964883267584", SheetName: "立直麻将(RCR)"},
		{ProductID: "1739484939113463808", SheetName: "四川麻将(SBR)"},
	}

	f := excelize.NewFile()
	defer f.Close()

	for _, product := range products {
		log.Printf("Fetching data for %s (productId: %s)...", product.SheetName, product.ProductID)

		players, err := fetchAllPlayers(product.ProductID)
		if err != nil {
			log.Fatalf("Failed to fetch data for %s: %v", product.SheetName, err)
		}

		log.Printf("Total players fetched for %s: %d", product.SheetName, len(players))

		if err := writeSheet(f, product.SheetName, players); err != nil {
			log.Fatalf("Failed to write sheet %s: %v", product.SheetName, err)
		}
	}

	// Set the first product sheet as active
	if index, err := f.GetSheetIndex(products[0].SheetName); err == nil {
		f.SetActiveSheet(index)
	}

	// Delete default Sheet1
	f.DeleteSheet("Sheet1")

	// Save file
	if err := f.SaveAs("mahjong_rank.xlsx"); err != nil {
		log.Fatalf("Failed to save Excel file: %v", err)
	}

	log.Println("Data saved to mahjong_rank.xlsx")
}
