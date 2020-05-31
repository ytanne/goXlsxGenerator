package goxlsx

import (
	"fmt"
	"log"
	"regexp"
	"strconv"
	//"sort"

	"github.com/360EntSecGroup-Skylar/excelize"
)

const (
	sheet = "Sheet1"
)

type cell struct {
	data string
}

/*InitExcel creates spreadsheets and declares all namings */
func InitExcel() *excelize.File {
	f := excelize.NewFile()

	title := fmt.Sprintf("Reputation result")
	f.MergeCell(sheet, "A1", "D1")
	f.SetCellValue(sheet, "A1", title)

	headers := []string{
		"URL", "IP", "IP Status",
		"Reputation",
	}
	fillRowHead(headers, f, "A2")

	return f
}

func fillRowHead(cells []string, f *excelize.File, sp string) {
	var c cell

	//sp is starting point
	c.init(sp)
	for _, cc := range cells {
		f.SetCellValue(sheet, c.data, cc)
		c.moveRight()
	}
}

func fillRowResult(result urlRep, f *excelize.File, sp string) {
	var c cell
	var initRow cell

	c.init(sp)
	initRow = c
	f.MergeCell(sheet, c.data, c.moveDown(len(result.IP) - 1))
	f.SetCellValue(sheet, c.data, result.URL)
	c.moveRight()

	reputations := result.Resp.Results

	for _, r := range reputations {
		initRow = c
		f.SetCellValue(sheet, c.data, r.IP)
		c.moveRight()
		f.SetCellValue(sheet, c.data, r.Queries.Info.IPstatus)
		c.moveRight()
		f.SetCellValue(sheet, c.data, r.Queries.Info.Reputation)
		c = initRow
		c.data = c.moveDown(1)
	}
}

/*CloseExcel creates excel file with provided name */
func CloseExcel(f *excelize.File, name string) error {
	if err := f.SaveAs(name + ".xlsx"); err != nil {
		log.Printf("Couldn't create excel file %v\n", name)
		log.Println(err)
		return err
	}
	return nil
}

func (c *cell) init(id string) string {
	c.data = id
	return c.data
}

func (c *cell) moveRight() {
	alre := regexp.MustCompile(`^[A-Z]+`)
	numre := regexp.MustCompile(`[\d]+$`)
	foundAlphaByte := alre.FindString(c.data)
	foundNumStr := numre.FindString(c.data)
	foundAlphaByte = incrementCell(foundAlphaByte)
	c.data = foundAlphaByte + foundNumStr
}

func (c *cell) moveDown(steps int) string {
	alre := regexp.MustCompile(`^[A-Z]+`)
	numre := regexp.MustCompile(`[\d]+$`)
	foundAlphaByte := alre.FindString(c.data)
	foundNumStr := numre.FindString(c.data)
	numstr, _ := strconv.Atoi(foundNumStr)
	foundNumStr = strconv.Itoa(numstr + steps)
	return foundAlphaByte + foundNumStr
}

func incrementCell(cell string) string {
	var counter int
	l := len(cell) - 1
	newStr := []rune(cell)

	for l >= 0 {
		if int(newStr[l]) == 90 {
			newStr[l] = 'A'
			counter++
		} else {
			alpha := int(newStr[l])
			alpha++
			newStr[l] = rune(alpha)
			break
		}
		l--
	}
	if counter >= 1 && l == -1 {
		return "A" + string(newStr)
	}
	return string(newStr)
}
