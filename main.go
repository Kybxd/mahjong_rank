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

func fetchPage(pageIndex int) ([]Player, error) {
	var reply Reply
	_, err := requests.Get(
		"https://wxapi.mahjonget.com/cert/open/score/page",
		requests.ParamPairs(
			"productId", "1739485995675418624",
			"pageIndex", strconv.Itoa(pageIndex),
			"pageSize", "100",
		),
		requests.ToJSON(&reply),
	)
	if err != nil {
		return nil, err
	}

	if !reply.Success {
		return nil, fmt.Errorf("API error: %s", reply.Msg)
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
	return players, nil
}

func saveToExcel(players []Player, filename string) error {
	f := excelize.NewFile()
	defer f.Close()

	// Create a new sheet
	index, err := f.NewSheet("Players")
	if err != nil {
		return err
	}

	// Set headers
	headers := []string{"Ranking", "Name", "UID", "CertName", "Score"}
	for col, header := range headers {
		cell, _ := excelize.CoordinatesToCellName(col+1, 1)
		f.SetCellValue("Players", cell, header)
	}

	// Fill data
	for row, player := range players {
		f.SetCellValue("Players", fmt.Sprintf("A%d", row+2), player.Ranking)
		f.SetCellValue("Players", fmt.Sprintf("B%d", row+2), player.Name)
		f.SetCellValue("Players", fmt.Sprintf("C%d", row+2), player.UID)
		f.SetCellValue("Players", fmt.Sprintf("D%d", row+2), player.CertName)
		f.SetCellValue("Players", fmt.Sprintf("E%d", row+2), player.Score)
	}

	// Set active sheet
	f.SetActiveSheet(index)

	// Delete Sheet1
	f.DeleteSheet("Sheet1")

	// Save file
	if err := f.SaveAs(filename); err != nil {
		return err
	}

	return nil
}

func main() {
	// Fetch first page to get total page count
	firstPagePlayers, err := fetchPage(1)
	if err != nil {
		log.Fatalf("Failed to fetch first page: %v", err)
	}

	// Get total page count from API
	var firstPageReply Reply
	_, err = requests.Get(
		"https://wxapi.mahjonget.com/cert/open/score/page",
		requests.ParamPairs(
			"productId", "1739485995675418624",
			"pageIndex", "1",
			"pageSize", "100",
		),
		requests.ToJSON(&firstPageReply),
	)
	if err != nil {
		log.Fatalf("Failed to get total page count: %v", err)
	}

	totalPages := firstPageReply.Data.TotalPageCount
	if totalPages <= 0 {
		totalPages = 1
	}

	// Collect all players
	var allPlayers []Player
	allPlayers = append(allPlayers, firstPagePlayers...)

	log.Printf("Total pages: %d", totalPages)
	log.Printf("Fetched page 1, players: %d", len(firstPagePlayers))

	// Fetch remaining pages
	for page := 2; page <= totalPages; page++ {
		players, err := fetchPage(page)
		if err != nil {
			log.Printf("Warning: Failed to fetch page %d: %v", page, err)
			continue
		}
		allPlayers = append(allPlayers, players...)
		log.Printf("Fetched page %d, players: %d", page, len(players))
	}

	log.Printf("Total players fetched: %d", len(allPlayers))

	// Save to Excel
	err = saveToExcel(allPlayers, "mahjong_rank.xlsx")
	if err != nil {
		log.Fatalf("Failed to save Excel file: %v", err)
	}

	log.Println("Data saved to mahjong_rank.xlsx")
}
