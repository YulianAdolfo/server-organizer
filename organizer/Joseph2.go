// code developed by  Yulian Adolfo Rojas - yulinarojas2000@gmail.com

package main

import (
	"fmt"
	"math"
	"strconv"
	"strings"

	"github.com/xuri/excelize/v2" // download this library - this library we use to work with .xlsx files
)

func findReferencePoint(pathFile, sheetName, booknameToSave string) {
	// Opening the excel book
	var electronicBill *excelize.File = openFileExcelFact(pathFile)
	// Getting the excel book columns
	electronicBillColumns, err := electronicBill.GetCols("Sheet0")
	if err != nil {
		panic("Ha ocurrido un error mientras se intentaba leer el excel")
	}
	// Go through all excel book looking for the word 'NUMERO_DOCUMENTO', 'PRE_ID_PROVEEDOR...'
	positionColumn := 0
	indexStartOne := 0
	indexEndOne := 0
	indexStartTwo := 0
	indexEndTwo := 0
	indexToSearch := [4]string{"NUMERO_DOCUMENTO", "PRE_ID_PROVEEDOR", "NUMERO_LIN", "RETENCIONES_ADICIONALES"}
	// Getting the indices
	for i := 0; i < len(indexToSearch); i++ {
		for electronicBillColumns[positionColumn][0] != indexToSearch[i] {
			positionColumn++
		}
		switch i {
		case 0:
			indexStartOne = positionColumn
		case 1:
			indexEndOne = positionColumn
		case 2:
			indexStartTwo = positionColumn
		case 3:
			indexEndTwo = positionColumn
		}
	}
	// indexStartOne and indexEndOne make reference to the header (IFCO, IFCR)
	// indexStartTwo and indexEndTwo make reference to the content of those headers
	organizeDataExcel(indexStartOne, indexEndOne, indexStartTwo, indexEndTwo, pathFile, sheetName, booknameToSave)
}

func organizeDataExcel(index1, index2, index3, index4 int, pathF, sheetName, booknameToSave string) {
	// Opening the excel fact file
	var excelFactHosvital *excelize.File = openFileExcelFact(pathF)
	// Opening the excel fact file to upload
	var finalFactUpload *excelize.File = openFileExcelFactToUpload(booknameToSave)

	// Getting alphabet cell
	var cell []string = combinationAlphabetColumns()
	// Getting the length colums
	var lenExcelFactHosvitalRow, err1 = excelFactHosvital.GetRows("Sheet0")
	if err1 != nil {
		panic("error")
	}
	var lenExcelFactHosvitalCols, err2 = excelFactHosvital.GetCols("Sheet0")
	if err2 != nil {
		panic("error")
	}
	var extractedData string
	var stopLoop bool
	var positionAlphabetArray int = 1
	var complementAlphabet int = 3
	const headerNumber = "1.1"
	const subHeaderNumber = "2.1"
	// Opening the new excel fact file
	// here goes
	for i := 1; i < len(lenExcelFactHosvitalRow); i++ {
		// getting the value to analize
		valueToAnalize := lenExcelFactHosvitalCols[index1][i]
		// going through the first space in file excel
		for j := index1; j <= index2; j++ {
			extractedData += "[" + lenExcelFactHosvitalCols[j][i] + "]"
			n := strconv.Itoa(complementAlphabet)
			finalFactUpload.SetCellValue(sheetName, cell[0]+n, headerNumber)
			finalFactUpload.SetCellValue(sheetName, cell[positionAlphabetArray]+n, lenExcelFactHosvitalCols[j][i])
			positionAlphabetArray++
		}
		complementAlphabet++
		positionAlphabetArray = 1
		fmt.Println(extractedData)
		extractedData = ""
		fmt.Println("se aumenta un fila en el otro excel")
		for k := index3; k <= index4; k++ {
			extractedData += "[" + lenExcelFactHosvitalCols[k][i] + "]"
			n := strconv.Itoa(complementAlphabet)
			finalFactUpload.SetCellValue(sheetName, cell[0]+n, subHeaderNumber)
			finalFactUpload.SetCellValue(sheetName, cell[positionAlphabetArray]+n, lenExcelFactHosvitalCols[k][i])
			positionAlphabetArray++
		}
		fmt.Println(extractedData)
		complementAlphabet++
		positionAlphabetArray = 1
		extractedData = ""
		for !stopLoop {
			if len(lenExcelFactHosvitalRow) == i+1 {
				fmt.Println("se ha sobre pasado")
				break
			}
			if valueToAnalize == lenExcelFactHosvitalCols[index1][i+1] {
				for k := index3; k <= index4; k++ {
					extractedData += "[" + lenExcelFactHosvitalCols[k][i+1] + "]"
					n := strconv.Itoa(complementAlphabet)
					finalFactUpload.SetCellValue(sheetName, cell[0]+n, subHeaderNumber)
					finalFactUpload.SetCellValue(sheetName, cell[positionAlphabetArray]+n, lenExcelFactHosvitalCols[k][i+1])
					positionAlphabetArray++
				}
				fmt.Println(extractedData)
				extractedData = ""
				positionAlphabetArray = 1
				complementAlphabet++
				i++
			} else {
				stopLoop = true
				positionAlphabetArray = 1
			}
		}
		stopLoop = false
	}
	err := finalFactUpload.SaveAs(booknameToSave + ".xlsx")
	if err != nil {
		panic("ha ocurrido un error")
	}
	startFix(finalFactUpload, sheetName, pathF, booknameToSave)
}

func openFileExcelFact(pathF string) *excelize.File {
	fmt.Println(pathF + " ruta")
	file, err := excelize.OpenFile(pathF)
	if err != nil {
		panic("hubo un error 1")
	}
	return file
}
func openFileExcelFactToUpload(booknameToSave string) *excelize.File {
	file, err := excelize.OpenFile(booknameToSave + ".xlsx")
	if err != nil {
		panic("hubo un error 2")
	}
	return file
}
func excelStructure(sheetName, booknameToSave string) {
	// Getting the headers
	var superiorHeader, inferiorHeader []string = getColumnsValues()
	var quantityColumns []string = combinationAlphabetColumns()
	var setWidthColum []float64 = getColumnsDimension()
	var excelFile *excelize.File = createExcelFile(booknameToSave)
	var n, n2 string = "1", "2"

	for i := 0; i < len(quantityColumns); i++ {
		excelFile.SetCellValue(sheetName, quantityColumns[i]+n, superiorHeader[i])

		if i < len(inferiorHeader) {
			excelFile.SetCellValue(sheetName, quantityColumns[i]+n2, inferiorHeader[i])
		}
		excelFile.SetColWidth(sheetName, quantityColumns[i], quantityColumns[i], setWidthColum[i])
	}

	if err := excelFile.SaveAs(booknameToSave + ".xlsx"); err != nil {
		panic("hubo un error al guardar el libro")
	}
	fmt.Println("agregada las lineas")
}
func createExcelFile(booknameToSave string) *excelize.File {
	excelFile := excelize.NewFile()
	excelFile.SaveAs(booknameToSave + ".xlsx")
	return excelFile
}
func getColumnsDimension() []float64 {
	var getDimensionCol []string
	var getWidthCol []float64
	var widthExcel float64
	dimension, _ := excelize.OpenFile("../modelFormat/width-height.xlsx")
	getDimensionCol = combinationAlphabetColumns()
	for i := 0; i < len(getDimensionCol); i++ {
		widthExcel, _ = dimension.GetColWidth("getCol", getDimensionCol[i])
		getWidthCol = append(getWidthCol, math.Floor(widthExcel*100)/100)
	}
	return getWidthCol
}
func getColumnsValues() ([]string, []string) {
	headerFields := []string{"1.0", "NUMERO_DOCUMENTO", "TIPO_DOCUMENTO", "SUBTIPO_DOCUMENTO", "TIPO_OPERACION",
		"DIVISA", "FECHA_DOCUMENTO", "REF_PEDIDO", "UNI_ORG", "FECHA_VENCIMIENTO",
		"MOTIVO_RECT", "INCOTERM", "DOCUMENTOS REFERENCIADOS", "PRE-ID_CLIENTE-DC",
		"TIPO_DOC_IDENTIDAD_CLIENTE", "REGIMEN_CLIENTE", "RAZON_SOCIAL_CLIENTE",
		"NOMBRE_CLIENTE", "APELLIDO1_CLIENTE", "APELLIDO2_CLIENTE", "TIPO_PERSONA_CLIENTE",
		"DIRECCION_CLIENTE", "AREA_CLIENTE", "CIUDAD_CLIENTE", "DISTRITO_CLIENTE", "CODIGO_POSTAL_CLIENTE",
		"TELEFONO_CLIENTE", "EMAIL_CLIENTE", "PAIS_CLIENTE", "MATRICULA_MERCANTIL", "CARACTERISTICAS_FISCALES",
		"TRIBUTOS", "ACTIVIDADES_ECONOMICAS", "IMPORTE", "PORCENTAJE_DESCUENTO", "DESCUENTO", "MOTIVO_DESCUENTO",
		"TOTAL_ANTICIPOS", "TOTAL", "APAGAR", "IMPUESTOS", "RETENCIONES", "DATOS ADICIONALES", "FORMA_PAGO",
		"MEDIO_PAGO", "ENTIDAD_BANCARIA", "NUMERO_CUENTA", "BENEFICIARIO", "FECHA_PAGO", "COD_DIVISA_ORIGEN", "COD_DIVISA_DESTINO",
		"TIPO_CAMBIO", "FECHA_TIPO_CAMBIO", "EMAIL_ENVIO", "Anticipos", "DIRECCION_FACTURA", "AREA_FACTURA", "CIUDAD_FACTURA",
		"CODIGO_POSTAL_FACTURA", "PAIS_FACTURA", "PRE-ID_PROVEEDOR-DC"}
	headerFields2 := []string{"2.0", "NUMERO_LIN", "DESCRIPCION_LIN", "ID_ESTANDAR_LIN", "CODIGO_ESTANDAR_LIN", "UNIDAD_MEDIDA_LIN",
		"MARCA_LIN", "MODELO_LIN", "PRE-ID_MANDANTE-DC", "UNIDADES_LIN", "PRECIO_UNIDAD_LIN", "IMPORTE_LIN",
		"PORCENTAJE_DESCUENTO_LIN", "DESCUENTO_LIN", "BASE_LIN", "PORCENTAJE_IMPUESTO_LIN", "IMPUESTO_LIN",
		"CODIGO_IMPUESTO_LIN_N", "DATOS ADICIONALES_ITEM", "IMPUESTOS_ADICIONALES", "RETENCIONES_ADICIONALES"}
	return headerFields, headerFields2
}
func combinationAlphabetColumns() []string {
	alphabetBase := []string{"A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M",
		"N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z"}
	cantIteration := len(alphabetBase)
	for i := 0; i < 2; i++ {
		if i == 1 {
			cantIteration = 9
		}
		for j := 0; j < cantIteration; j++ {
			alphabetBase = append(alphabetBase, alphabetBase[i]+alphabetBase[j])
		}
	}
	return alphabetBase
}
func fixingNumberConsecutive(fileToFix *excelize.File, sheetName string) {
	rows, err := fileToFix.GetRows("Sheet1")
	if err != nil {
		fmt.Println(err)
		panic("Ha ocurrido un error 1")
	}
	columns, err := fileToFix.GetCols("Sheet1")
	if err != nil {
		panic("Ha ocurrido un error 2")
	}
	var counter int = 0
	if err != nil {
		panic("erorr de estilo")
	}
	// defining cell styles
	textPosition, err := fileToFix.NewStyle(`{"alignment":{"horizontal":"left"}}`)
	if err != nil {
		fmt.Println("ha ocurrido un error al aplicar crear estilo")
	}

	for i := 0; i < len(rows); i++ {
		_, err := strconv.Atoi(columns[1][i])
		if err != nil {
			counter = 0
		} else {
			counter++
			n := strconv.Itoa(i + 1)
			cell := "B" + n
			fileToFix.SetCellValue(sheetName, cell, counter)
			fileToFix.SetCellStyle(sheetName, cell, cell, textPosition)
		}
	}
	analizingEmail(fileToFix, sheetName)

}
func saveBook(f *excelize.File) {
	err := f.Save()
	if err != nil {
		fmt.Println("Ha ocurrido un error al guardar el archivo")
	}
}
func analizingEmail(f *excelize.File, sheetName string) {
	var uncorrectPosibleEmails []string
	var uncorrectPositionEmails []string
	columnExcel, err := f.GetCols(sheetName)
	if err != nil {
		fmt.Println(err)
		fmt.Println("ha ocurrido un error")
	}
	rowExcel, err := f.GetRows(sheetName)
	if err != nil {
		fmt.Println(err)
		fmt.Println("ha ocurrido un error")
	}
	searchedField := "EMAIL_CLIENTE"
	i := 0
	for columnExcel[i][0] != searchedField {
		i++
	}
	fmt.Println("longitud row excel: ", len(rowExcel))
	fmt.Println("encontrado en posicion: ", i, columnExcel[i][0])
	iteration := len(rowExcel) - 1
	for j := 2; j < iteration; j++ {
		fmt.Println(columnExcel[i][j])
		if columnExcel[i][j] != "" {

			if strings.Contains(strings.ToLower(columnExcel[i][j]), "@gmail.com") ||
				strings.Contains(strings.ToLower(columnExcel[i][j]), "@outlook.com") ||
				strings.Contains(strings.ToLower(columnExcel[i][j]), "@hotmail.com") ||
				strings.Contains(strings.ToLower(columnExcel[i][j]), "@outlook.es") ||
				strings.Contains(strings.ToLower(columnExcel[i][j]), "@yahoo.com") ||
				strings.Contains(strings.ToLower(columnExcel[i][j]), "@aol.com") ||
				strings.Contains(strings.ToLower(columnExcel[i][j]), "@misena.edu.co") ||
				strings.Contains(strings.ToLower(columnExcel[i][j]), "@zohomail.com") ||
				strings.Contains(strings.ToLower(columnExcel[i][j]), "@ucaldas.edu.co") {

				deleteBlankSpaces := strings.Replace(columnExcel[i][j], " ", "", -1)
				lenghtEmail := len(deleteBlankSpaces)
				convertRune := []rune(deleteBlankSpaces)
				for string(convertRune[lenghtEmail-1]) == "." {
					lenghtEmail--
				}
				if lenghtEmail == len(deleteBlankSpaces) {
					continue
				} else {
					n := strconv.Itoa(j + 1)
					cell := "AB" + n
					f.SetCellValue(sheetName, cell, string(convertRune[0:lenghtEmail]))
					fmt.Println(string(convertRune[0:lenghtEmail]), "posicionado en ", j)
				}
			} else {
				uncorrectPosibleEmails = append(uncorrectPosibleEmails, columnExcel[i][j])
				p := strconv.Itoa(j + 1)
				uncorrectPositionEmails = append(uncorrectPositionEmails, p)
			}
		}
	}
	fmt.Println("posibles correos incorrectos")
	for i := 0; i < len(uncorrectPosibleEmails); i++ {
		fmt.Println(uncorrectPosibleEmails[i])
		fmt.Println("en la posicion")
		fmt.Println(uncorrectPositionEmails[i])
	}
	saveBook(f)
}
func startFix(f *excelize.File, sheetName, pathWay, bookName string) {
	file, err := excelize.OpenFile("../file-excel/FACTURACION ELECTRONICA 05-02-2020.xlsx")
	if err != nil {
		panic("Hubo un error al abrir el archivo")
	}
	fixingNumberConsecutive(file, sheetName)
}

func startBuild(pathway, sheetname, booknameToSave string) {
	excelStructure(sheetname, booknameToSave)
	getColumnsDimension()
	findReferencePoint(pathway, sheetname, booknameToSave)
}
func main() {
	/* // requesting the file route to the user
	setPathRoute := bufio.NewReader(os.Stdin)
	fmt.Print("file route with filename (.xlsx) --> ")
	getPathRoute, err := setPathRoute.ReadString('\n')

	if err != nil {
		panic("Ha ocurrido un error en la ruta de archivo")
	}
	getPathRoute = strings.TrimSpace(getPathRoute)

	getPathRoute = strings.Replace(getPathRoute, "\\", "/", -1) */
	name := "../file-excel/FACTURACION ELECTRONICA 05-02-2020"
	pathWay := name                          // route where is the file that we'll sort
	sheetName := "Sheet1"                    // sheet name where is the content to sort
	bookName := name + "-srt"                // the new file name sorted
	startBuild(pathWay, sheetName, bookName) // beginning to sort

}
