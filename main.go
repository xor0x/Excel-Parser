package main

import (
	"bufio"
	"io/ioutil"
	"os"
	"time"

	"fyne.io/fyne/v2"
	"fyne.io/fyne/v2/app"
	"fyne.io/fyne/v2/container"
	"fyne.io/fyne/v2/dialog"
	"fyne.io/fyne/v2/widget"
)

type config struct {
	startDate *widget.Entry
	endDate *widget.Entry
	loadButton *widget.Button
	saveButton *widget.Button
	loadEntry *widget.Entry
	saveEntry *widget.Entry
	startButton *widget.Button
	progressBar *widget.ProgressBar 
	textDone *widget.Label
	sheetList *widget.Label
}

var cnf config


func main() {
	
	a := app.New()

	win := a.NewWindow("Excel parser")
	r, _ := fyne.LoadResourceFromPath("Icon.png")
    // change icon
    a.SetIcon(r)
	
	progressBar, textDone, fromButton, toButton, sheetLabel:= cnf.MakeUI(win)
	loadButton, startButton, saveButton := cnf.Buttons(win)
	startDate, endDate, loadInput, saveInput := cnf.Entries(win)

	
	win.SetContent(container.NewWithoutLayout(
		loadButton,
		startDate,
		loadInput,
		saveInput,
		endDate,
		startButton,
		progressBar,
		textDone,
		saveButton,
		fromButton,
		toButton,
		sheetLabel,
		),)
	win.Resize(fyne.NewSize(600, 600))
	// win.SetFixedSize(true)
	win.CenterOnScreen()
	win.ShowAndRun()
}


func (app *config) MakeUI(win fyne.Window) (*widget.ProgressBar,*widget.Label, *widget.Button, *widget.Button, *widget.Label) {
	
	w := win.Canvas()
	iconFile, _ := os.Open("calendar.jpg")
	r := bufio.NewReader(iconFile)

	b, _ := ioutil.ReadAll(r)
	fromButton := widget.NewButtonWithIcon("", fyne.NewStaticResource("icon", b), func() {
		NewCalendarAtPos(w, fyne.NewPos(10, 10),time.Now(), func(t time.Time){
			app.startDate.SetText(t.Format("01/02/2006"))
		})
	})

	fromButton.Resize(fyne.NewSize(38, 38))
	fromButton.Move(fyne.NewPos(260, 51))

	toButton := widget.NewButtonWithIcon("", fyne.NewStaticResource("icon", b), func() {
		NewCalendarAtPos(w, fyne.NewPos(300, 10),time.Now(), func(t time.Time){
			app.endDate.SetText(t.Format("01/02/2006"))
		})
	})
	toButton.Resize(fyne.NewSize(38, 38))
	toButton.Move(fyne.NewPos(550, 51))


	sheetsLabel := widget.NewLabel("Sheets: 0/0")
	sheetsLabel.Move(fyne.NewPos(10, 270))
	sheetsLabel.TextStyle.Bold = true
	app.sheetList = sheetsLabel
	

	progressBar := widget.NewProgressBar()
	progressBar.Resize(fyne.NewSize(500, 38)) // my widget size
	progressBar.Move(fyne.NewPos(10, 300))     // position of widget
	app.progressBar = progressBar


	textDone := widget.NewLabel("DONE")
	textDone.Move(fyne.NewPos(510, 300))
	textDone.TextStyle.Bold = true
	textDone.Hidden = true
	app.textDone = textDone

	return  progressBar, textDone, fromButton, toButton, sheetsLabel
}

func (app *config) Entries(win fyne.Window) (*widget.Entry, *widget.Entry, *widget.Entry, *widget.Entry){
	
	startDate := widget.NewEntry()
    startDate.SetPlaceHolder("mm/dd/yyyy")
	startDate.Resize(fyne.NewSize(250, 38)) // my widget size
    startDate.Move(fyne.NewPos(10, 50))     // position of widget


	endDate := widget.NewEntry()
    endDate.SetPlaceHolder("mm/dd/yyyy")
	endDate.Resize(fyne.NewSize(250, 38)) // my widget size
    endDate.Move(fyne.NewPos(300, 50))     // position of widget
	app.startDate = startDate
	app.endDate = endDate


	loadInput := widget.NewEntry()
	loadInput.SetPlaceHolder("")
	loadInput.Resize(fyne.NewSize(250, 38)) // my widget size
	loadInput.Move(fyne.NewPos(10, 100))     // position of widget
	app.loadEntry = loadInput


	saveInput := widget.NewEntry()
	saveInput.SetPlaceHolder("")
	saveInput.Resize(fyne.NewSize(250, 38)) // my widget size
	saveInput.Move(fyne.NewPos(10, 150))     // position of widget
	app.saveEntry = saveInput

	return startDate, endDate, loadInput, saveInput
}


func (app *config) Buttons(win fyne.Window) (*widget.Button, *widget.Button, *widget.Button){
	loadButton := widget.NewButton("Load excel file", app.openFunc(win))
	loadButton.Resize(fyne.NewSize(170, 38))
	loadButton.Move(fyne.NewPos(270, 100))
	app.loadButton = loadButton

	saveButton := widget.NewButton("Save file", app.saveFunc(win))
	saveButton.Resize(fyne.NewSize(170, 38))
	saveButton.Move(fyne.NewPos(270, 150))
	app.saveButton = saveButton


	startButton := widget.NewButton("Start", app.startFunc(win))
	startButton.Resize(fyne.NewSize(200, 50))
	startButton.Move(fyne.NewPos(10, 200))
	app.startButton = startButton

	return loadButton, startButton, saveButton
}


func (app *config) openFunc(win fyne.Window) func(){
	return func(){
		file_Dialog := dialog.NewFileOpen(
			func(read fyne.URIReadCloser, err error) {
				if err != nil {
					dialog.ShowError(err, win)
					return
				}
				if read == nil{
					return
				}
				defer read.Close()

				app.loadEntry.SetText(read.URI().Path())
				app.loadSheets()
			}, win)
		file_Dialog.Resize(fyne.NewSize(500, 500))
		file_Dialog.Show()
		// Show file selection dialog.
	}
}


func (app *config) saveFunc(win fyne.Window) func(){
	return func(){
		saveDialog := dialog.NewFileSave(func(write fyne.URIWriteCloser, err error){
			if err != nil{
				dialog.ShowError(err, win)
				return
			}
			
			if write == nil {
				// user cancelled
				return
			}

			app.saveEntry.SetText(write.URI().Path())
			
		}, win)
		saveDialog.SetFileName("NewFile.xlsx")
		saveDialog.Resize(fyne.NewSize(500, 500))
		saveDialog.Show()
	}
}


func (app *config) startFunc(win fyne.Window) func(){
	return app.getData
}


