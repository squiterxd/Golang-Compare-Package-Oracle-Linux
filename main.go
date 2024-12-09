package main

import (
	"fmt"
	"log"

	"github.com/xuri/excelize/v2"
)

type PackageInfo struct {
	Name     string
	Version  string          // เพิ่มเวอร์ชันในโครงสร้าง
	Machines map[string]bool // เครื่อง: มีหรือไม่มี
}

func main() {
	// เปิดไฟล์ Excel ที่มีหลายแท็บ
	f, err := excelize.OpenFile("ListPackage.xlsx")
	if err != nil {
		log.Fatal(err)
	}

	// รายชื่อเครื่อง
	machines := []string{"ora19cprd", "ora19cstd", "ora19crpt", "ora19cdev", "ora19cdr"}

	// สร้างแผนที่เพื่อจัดเก็บข้อมูลแพ็กเกจ
	packageMap := make(map[string]*PackageInfo)

	// ดึงชื่อแท็บทั้งหมด
	sheets := f.GetSheetList()

	// อ่านข้อมูลจากแต่ละแท็บ
	for _, machine := range machines {
		found := false
		for _, sheet := range sheets {
			if sheet == machine {
				found = true
				break
			}
		}

		// ถ้าไม่พบแท็บให้ข้ามไป
		if !found {
			log.Printf("แท็บ %s ไม่พบในไฟล์ Excel", machine)
			continue
		}

		// อ่านข้อมูลจากแท็บ
		rows, err := f.GetRows(machine)
		if err != nil {
			log.Fatal(err)
		}

		for _, row := range rows {
			if len(row) > 1 { // ตรวจสอบว่ามีข้อมูลในสองคอลัมน์ขึ้นไป
				packageName := row[0] // สมมุติว่าแพ็กเกจอยู่ในคอลัมน์แรก
				version := row[1]     // สมมุติว่าเวอร์ชันอยู่ในคอลัมน์ที่สอง
				if _, exists := packageMap[packageName]; !exists {
					packageMap[packageName] = &PackageInfo{Name: packageName, Version: version, Machines: make(map[string]bool)}
				}
				packageMap[packageName].Machines[machine] = true // เครื่องนี้มีแพ็กเกจนี้
			}
		}
	}

	// สร้างไฟล์ Excel ใหม่
	newFile := excelize.NewFile()
	newSheet := "Package Status"
	newFile.NewSheet(newSheet)

	// เขียนหัวข้อ
	newFile.SetCellValue(newSheet, "A1", "Package Name")
	for i, machine := range machines {
		newFile.SetCellValue(newSheet, fmt.Sprintf("%s1", string('B'+i)), machine) // เขียนชื่อเครื่อง
	}

	// เขียนข้อมูลแพ็กเกจและสถานะ
	row := 2
	for _, pkg := range packageMap {
		newFile.SetCellValue(newSheet, fmt.Sprintf("A%d", row), pkg.Name) // ชื่อแพ็กเกจ
		for i, machine := range machines {
			if _, exists := pkg.Machines[machine]; exists {
				newFile.SetCellValue(newSheet, fmt.Sprintf("%s%d", string('B'+i), row), pkg.Version) // แสดงเวอร์ชัน
			} else {
				newFile.SetCellValue(newSheet, fmt.Sprintf("%s%d", string('B'+i), row), "Not Install") // ถ้าไม่มี
			}
		}
		row++
	}

	// บันทึกไฟล์ Excel ใหม่
	if err := newFile.SaveAs("package_status.xlsx"); err != nil {
		log.Fatal(err)
	}

	fmt.Println("สร้างไฟล์ Excel ใหม่เสร็จสิ้น: package_status.xlsx")
}
