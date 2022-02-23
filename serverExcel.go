package main

import (
	"fmt"
	"html/template"
	"io/ioutil"
	"log"
	"math"
	"net"
	"net/http"
	"strconv"
	"strings"

	"github.com/xuri/excelize/v2" // download this library - this library we use to work with .xlsx files
)

// FileExcel is a struct
type FileExel struct {
	Filename, LinkFile, AddrServer string
}

var fileHTMLServer = template.Must(template.ParseFiles("index.html", "download.html"))

func upRunServer(w http.ResponseWriter, r *http.Request) {
	fileHTMLServer.ExecuteTemplate(w, "index.html", nil)
}
func fileSentRequest(w http.ResponseWriter, r *http.Request) {
	var filen string
	if r.Method == "POST" {
		// parsing the form enctype="multipart/form-data"
		r.ParseMultipartForm(10 << 20) // specified the size of the file. max size is ten to twenty megabytes

		// retrieve file
		// FormFile receives the name of the tag file
		// this return three values: 1- the file - 2-characteristic 3- err

		fileContent, cFile, err := r.FormFile("file-excel")
		if err != nil {
			fmt.Fprintf(w, "an error has occurred")
		}
		// creating a temporal file
		// we use the library io/ioutil
		// method: tempFile() --> receives: 1- the path, 2- a pattern text

		temporalFile, err := ioutil.TempFile("file-excel", "file_to_analize-*.xlsx")
		if err != nil {
			fmt.Println("File error")
		}
		// read the file in bytes

		informationInBytesOfSentFile, err := ioutil.ReadAll(fileContent)
		if err != nil {
			fmt.Println("Error file 2")
		}
		// writing the bytes
		temporalFile.Write(informationInBytesOfSentFile)
		defer temporalFile.Close()
		defer fileContent.Close()

		var nameF string = cFile.Filename
		runes := []rune(nameF)
		nameF = string(runes[0 : len(nameF)-5])
		//fmt.Println("file is uploaded to the server")
		p := strings.Replace(temporalFile.Name(), "\\", "/", 1)

		startBuild(p, "Sheet1", nameF) // method to start to organize the excel file
		filen = nameF + ".xlsx"
	}
	pathDownload := "ready-file/" + filen
	defer downloadPageFile(w, filen, pathDownload)
}
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
		extractedData = ""
		for k := index3; k <= index4; k++ {
			extractedData += "[" + lenExcelFactHosvitalCols[k][i] + "]"
			n := strconv.Itoa(complementAlphabet)
			finalFactUpload.SetCellValue(sheetName, cell[0]+n, subHeaderNumber)
			finalFactUpload.SetCellValue(sheetName, cell[positionAlphabetArray]+n, lenExcelFactHosvitalCols[k][i])
			positionAlphabetArray++
		}
		complementAlphabet++
		positionAlphabetArray = 1
		extractedData = ""
		for !stopLoop {
			if len(lenExcelFactHosvitalRow) == i+1 {
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
	err := finalFactUpload.SaveAs("ready-file/" + booknameToSave + ".xlsx")
	if err != nil {
		panic("ha ocurrido un error")
	}
	startFix(finalFactUpload, sheetName, pathF, booknameToSave)
}

func openFileExcelFact(pathF string) *excelize.File {
	file, err := excelize.OpenFile(pathF)
	if err != nil {
		panic("hubo un error")
	}
	return file
}
func openFileExcelFactToUpload(booknameToSave string) *excelize.File {
	file, err := excelize.OpenFile("ready-file/" + booknameToSave + ".xlsx")
	if err != nil {
		panic("hubo un error")
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

	if err := excelFile.SaveAs("ready-file/" + booknameToSave + ".xlsx"); err != nil {
		panic("hubo un error al guardar el libro")
	}
}
func createExcelFile(booknameToSave string) *excelize.File {
	excelFile := excelize.NewFile()
	excelFile.SaveAs("ready-file/" + booknameToSave + ".xlsx")
	return excelFile
}
func getColumnsDimension() []float64 {
	var getDimensionCol []string
	var getWidthCol []float64
	var widthExcel float64
	dimension, _ := excelize.OpenFile("modelFormat/width-height.xlsx")
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
		"CODIGO_POSTAL_FACTURA", "PAIS_FACTURA", "PRE-ID_PROVEEDOR-DC", "TIPO DE OPERACIÓN", "CODIGO_PRESTADOR_SERVICIO",
		"TIPO_DOC_USUARIO", "NUMERO_DOC_USUARIO", "PRIMER_APELLIDO_USUARIO", "SEGUNDO_APELLIDO_USUARIO",
		"PRIMER_NOMBRE_USUARIO", "SEGUNDO_NOMBRE_USUARIO", "TIPO_USUARIO", "MODALIDADES_CONTRATACION_PAGO",
		"COBERTURA_PLAN_BENEFICIOS", "NUMERO_AUTORIZACION", "NUMERO_PREESCRIPCION_MIPRES", "NUMERO_IDENTIFICACION_MIPRES",
		"NUMERO_CONTRATO", "NUMERO_POLIZA", "FECHA_INICIO_FACTURACION", "FECHA_FIN_FACTURACION", "COPAGO", "CUOTA_MODERADORA",
		"CUOTA_RECUPERACION", "PAGOS_COMPARTIDOS", "NUMERO_DOCUMENTO_REF",
	}
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
	for i := 0; i < 3; i++ {
		if i == 1 {
			cantIteration = 7
		}
		for j := 0; j < cantIteration; j++ {
			alphabetBase = append(alphabetBase, alphabetBase[i]+alphabetBase[j])
		}
	}
	fmt.Println(alphabetBase)
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
	iteration := len(rowExcel) - 1
	for j := 2; j < iteration; j++ {
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
				}
			} else {
				uncorrectPosibleEmails = append(uncorrectPosibleEmails, columnExcel[i][j])
				p := strconv.Itoa(j + 1)
				uncorrectPositionEmails = append(uncorrectPositionEmails, p)
			}
		}
	}
	/* 	for i := 0; i < len(uncorrectPosibleEmails); i++ {
		fmt.Println(uncorrectPosibleEmails[i])
		fmt.Println("en la posicion")
		fmt.Println(uncorrectPositionEmails[i])
	} */
	saveBook(f)
}
func startFix(f *excelize.File, sheetName, pathWay, bookName string) {
	file, err := excelize.OpenFile("ready-file/" + bookName + ".xlsx")

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
func saveBook(f *excelize.File) {
	err := f.Save()
	if err != nil {
		fmt.Println("Ha ocurrido un error al guardar el archivo")
	}

}
func downloadPageFile(w http.ResponseWriter, filename, pathD string) {
	ip := sendIPServer().String()
	data := FileExel{
		Filename:   filename,
		LinkFile:   pathD,
		AddrServer: ip,
	}
	fileHTMLServer.ExecuteTemplate(w, "download.html", data)
}
func sendIPServer() net.IP {
	connection, err := net.Dial("udp", "8.8.8.8:80")
	if err != nil {
		log.Fatal(err)
	}
	defer connection.Close()

	ipServer := connection.LocalAddr().(*net.UDPAddr)

	return ipServer.IP
}
func main() {
	fmt.Println("please, go to http:" + sendIPServer().String() + ":2021 and enjoy this app!")
	http.Handle("/", http.FileServer(http.Dir("./")))
	http.HandleFunc("/file", upRunServer)
	http.HandleFunc("/uploadFile", fileSentRequest)
	http.ListenAndServe(":2021", nil)
}
