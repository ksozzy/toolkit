=begin
  #--Created by caidong<caidong@wps.cn> on 2018/2/12.
  #++Description:
=end
require 'sdk'
require 'json'
require 'Qt'
require_relative 'fileutils'

module AIOffice_Toolkit

  class Sample < KSO_SDK::JsBridge

    public
  
	
	#writer
    def w_open(filepath)
      if File.exist?(filepath)
        KSO_SDK.getApplication().Documents.Open(filepath)
        return true
      end
      return false
    end
	
	#presentation
    def p_open(filepath)
      if File.exist?(filepath)
        KSO_SDK.getApplication().Presentations.Open(filepath)
        return true
      end
      return false
    end
	
	#spreadsheet
    def s_open(filepath)
      if File.exist?(filepath)
        KSO_SDK.getApplication().Workbooks.Open(filepath)
        return true
      end
      return false
    end
	
    #以模板的形式打开Excel
    def s_openTemplate(filepath)
      KSO_SDK::Application.Workbooks.Add(filepath)
    end
	
    #保存Excel文档
    def s_save()
      KSO_SDK::Application.ActiveWorkbook.Save
    end
  
    #获取当前选中的单元格位置
    def s_getSelection()
      KSO_SDK::Application.Selection.Address
    end
	
    #选中
    def s_rangeSelect(range)
      KSO_SDK::Application.ActiveSheet.Range(range).Select
    end
	
    def s_rangeSelectInBook(book, sheet, range)
      KSO_SDK::Application.Workbooks(book).Activate
      KSO_SDK::Application.Sheets(sheet).Activate
	  KSO_SDK::Application.ActiveSheet.Range(range).Select
    end
	
    def s_getCellValue(cell)
      KSO_SDK::Application.ActiveSheet.Range(cell).Value
    end
	
	
    def s_getCellValueInBook(book, sheet, cell)
      KSO_SDK::Application.Workbooks(book).Sheets(sheet).Range(cell).Value
    end
	
    def s_getCellFormula(cell)
      KSO_SDK::Application.ActiveSheet.Range(cell).Formula
    end
    
    def s_getCellFormulaInBook(book, sheet, cell)
      KSO_SDK::Application.Workbooks(book).Sheets(sheet).Range(cell).Formula
    end
  
    def s_setRangeValue(range, value)
      KSO_SDK::Application.ActiveSheet.Range(range).Value = value
    end
	
    def s_setRangeValueInBook(book, sheet, range, value)
      KSO_SDK::Application.Workbooks(book).Sheets(sheet).Range(range).Value = value
    end
	
    def s_setCellFormula(range,value)
      KSO_SDK::Application.ActiveSheet.Range(range).Formula = value
    end
	
    def s_setCellFormulaInBook(book, sheet, range, value)
      KSO_SDK::Application.Workbooks(book).Sheets(sheet).Range(range).Formula = value
    end
	
    def s_getCellNumberFormat(range)
      KSO_SDK::Application.ActiveSheet.Range(range).NumberFormat
    end
	
    def s_getRangeNumberFormatInBook(book, sheet, range)
      KSO_SDK::Application.Workbooks(book).Sheets(sheet).Range(range).NumberFormat
    end
	
    def s_setRangeNumberFormat(range, format)
      KSO_SDK::Application.ActiveSheet.Range(range).NumberFormat = format
    end
	
    def s_setRangeNumberFormatInBook(book, sheet, range, format)
      KSO_SDK::Application.Workbooks(book).Sheets(sheet).Range(range).NumberFormat = format
    end
	
    def s_getCellColorInBook(book, sheet, range)
      KSO_SDK::Application.Workbooks(book).Sheets(sheet).Range(range).Interior.Color
    end
	
    def s_setCellColorInBook(book, sheet, range, color)
	  if (16777215 == color)
	    KSO_SDK::Application.Workbooks(book).Sheets(sheet).Range(range).Interior.Pattern = -4142
	  else
        KSO_SDK::Application.Workbooks(book).Sheets(sheet).Range(range).Interior.Color = color
	  end
    end
	
    def s_rangeCopy(range)
      KSO_SDK::Application.ActiveSheet.Range(range).Copy()
    end
	
    def s_rangeCopyInBook(book, sheet, range)
      KSO_SDK::Application.Workbooks(book).Sheets(sheet).Range(range).Copy()
    end
	
    def s_rangePaste(range)
      KSO_SDK::Application.ActiveSheet.Range(range).Select
      KSO_SDK::Application.ActiveSheet.Paste
    end
	
    def s_rangePasteInBook(book, sheet, range)
      KSO_SDK::Application.Workbooks(book).Sheets(sheet).Activate
      KSO_SDK::Application.ActiveSheet.Range(range).Select
      KSO_SDK::Application.ActiveSheet.Paste
    end
	
    def s_rangePasteValueAndFormat(range)
      KSO_SDK::Application.ActiveSheet.Range(range).Select
      KSO_SDK::Application.ActiveSheet.Range(range).PasteSpecial(13, -4142, false, false)
      KSO_SDK::Application.ActiveSheet.Range(range).PasteSpecial(-4163, -4142, false, false)
    end
	
    def s_setRangeTextToRight(range)
      KSO_SDK::Application.ActiveSheet.Range(range).HorizontalAlignment = -4152
    end
	
    def s_columnAutoFit(range)
      KSO_SDK::Application.Columns(range).EntireColumn.AutoFit
    end
	
    def s_setScreenUpdating(bUpdate)
      KSO_SDK::Application.ScreenUpdating = bUpdate
    end
	
	
	
	
	
	
    def deleteRange(range, shift)
      KSO_SDK::Application.ActiveSheet.Range(range).Delete(shift)
    end
	
	def insertRange(range)
		KSO_SDK::Application.ActiveSheet.Range(range).Insert()
	end
	def insertRow(range)
		KSO_SDK::Application.ActiveSheet.Range(range).Rows.Insert()
	end
	def insertCol(range)
		KSO_SDK::Application.ActiveSheet.Range(range).Columns.Insert()
	end
	
    def setAllRowsHidden(hidden)
      KSO_SDK::Application.ActiveSheet.Rows.Hidden = hidden
    end
    def setAllColsHidden(hidden)
      KSO_SDK::Application.ActiveSheet.Columns.Hidden = hidden
    end
    def setRangeRowsHidden(range, hidden)
      KSO_SDK::Application.ActiveSheet.Range(range).EntireRow.Hidden = hidden
    end
    def getRangeRowsHidden(range)
      KSO_SDK::Application.ActiveSheet.Range(range).Rows.Hidden
    end
	
    def setRangeColsHidden(range, hidden)
      KSO_SDK::Application.ActiveSheet.Range(range).EntireColumn.Hidden = hidden
    end
    def getRangeColsHidden(range)
      KSO_SDK::Application.ActiveSheet.Range(range).Columns.Hidden
    end
	
    def deleteRangeHyperlinks(range)
      KSO_SDK::Application.ActiveSheet.Range(range).Hyperlinks.Delete()
    end
	
    def deleteRangeComments(range)
      KSO_SDK::Application.ActiveSheet.Range(range).ClearComments()
    end
	
    def setRangeColor(range, color)
      KSO_SDK::Application.ActiveSheet.Range(range).Interior.ThemeColor = color
    end
	
	def hasMergerCells(range)
		KSO_SDK::Application.ActiveSheet.Range(range).MergeCells
	end
	def getMergerArea(range)
		KSO_SDK::Application.ActiveSheet.Range(range).MergeArea.Address
	end
	def rangeMerge(range)
		KSO_SDK::Application.ActiveSheet.Range(range).HorizontalAlignment = -4108
		KSO_SDK::Application.ActiveSheet.Range(range).VerticalAlignment = -4108
		KSO_SDK::Application.ActiveSheet.Range(range).WrapText = false
		KSO_SDK::Application.ActiveSheet.Range(range).Orientation = 0
		KSO_SDK::Application.ActiveSheet.Range(range).AddIndent = false
		KSO_SDK::Application.ActiveSheet.Range(range).IndentLevel = 0
		KSO_SDK::Application.ActiveSheet.Range(range).ShrinkToFit = false
		KSO_SDK::Application.ActiveSheet.Range(range).ReadingOrder = -5002
		KSO_SDK::Application.ActiveSheet.Range(range).Merge
	end
	def rangeUnMerge(range)
		KSO_SDK::Application.ActiveSheet.Range(range).HorizontalAlignment = 1
		KSO_SDK::Application.ActiveSheet.Range(range).VerticalAlignment = -4108
		KSO_SDK::Application.ActiveSheet.Range(range).WrapText = false
		KSO_SDK::Application.ActiveSheet.Range(range).Orientation = 0
		KSO_SDK::Application.ActiveSheet.Range(range).AddIndent = false
		KSO_SDK::Application.ActiveSheet.Range(range).IndentLevel = 0
		KSO_SDK::Application.ActiveSheet.Range(range).ShrinkToFit = false
		KSO_SDK::Application.ActiveSheet.Range(range).ReadingOrder = -5002
		KSO_SDK::Application.ActiveSheet.Range(range).UnMerge
	end
	
	
    def etFunc_isError(range)
      val = KSO_SDK::Application.WorksheetFunction.IsError(KSO_SDK::Application.ActiveSheet.Range(range))
	  val
    end
	
    def etFunc_isNumber(range)
      val = KSO_SDK::Application.WorksheetFunction.IsNumber(KSO_SDK::Application.ActiveSheet.Range(range))
	  val
    end
	
    #打开选择文件弹框
    def openFileDialog(title="打开文件",path= "C:", desc = "files", suffix="*.*")
      Qt::FileDialog::getOpenFileName(KSO_SDK::getCurrentMainWindow(), title,
        path,
        "#{desc} (#{suffix})")
    end
  
    #添加Sheet
    def addSheet()
      sheet = KSO_SDK::Application.WorkSheets.Add
      sheet.Name
    end
  
    #隐藏Sheet
    def hideSheet(name)
      KSO_SDK::Application.WorkSheets(name).Visible = false
    end
  
    #为单元格设置自动填充的内容
    def autoFill(src, sheet, dst)
      # to-do
    end
  
  
    #以模板的形式打开Word
    def openWordTemp(filepath)
      KSO_SDK::Application.Documents.Add(filepath)
    end
  
    #弹出Excel选择单元格选择窗
    def showInputBox_range(prompt, title)
      KSO_SDK::Application.InputBox(:prompt => prompt, :title => title, :type => 8).Address
    end
    def showInputBox_text(prompt, title)
      KSO_SDK::Application.InputBox(:prompt => prompt, :title => title, :type => 2)
    end
    def showInputBox_number(prompt, title)
      KSO_SDK::Application.InputBox(:prompt => prompt, :title => title, :type => 1)
    end
  
    #Excel文档另存为
    def excelSaveAs()
      filename = KSO_SDK::Application.GetSaveAsFilename()
      KSO_SDK::Application.ActiveWorkbook.SaveAs(filename)
    end
  
    #Excel关闭当前文档
    def closeActiveWorkbook()
      KSO_SDK::Application.ActiveWorkbook.Close()
    end
    
    #获取已使用的区域
    def getUsedRangeAddress()
      KSO_SDK::Application.ActiveSheet.UsedRange.Address
    end
  
    # 显示MessageBox
    def showMessageBox(title, text)
      btnMask = Qt::MessageBox::question(KSO_SDK.getCurrentMainWindow(), 'Title', 'ContentMessage', Qt::MessageBox::Yes, Qt::MessageBox::No)
    end
  
    #为单元格设置下来选值
    def setRangeInCellDropdownValidation(address, array)
      #{"type":3,"value":true,"alertStyle":1,"operator":1,"inCellDropdown":true,"formula1":"123,321,abc","formula2":""} 
      KSO_SDK::Application.ActiveSheet.Range(address).Validation().Add(3, 1, 1, array)
    end
  
    #为单元格添加批注
    def setComment(address, comment)
      KSO_SDK::Application.ActiveSheet.Range(address).AddComment(comment)
      nil
    end
	
	#获取文件全路径
	def getFullName()
	  klog val = KSO_SDK::Application.ActiveWorkbook.FullName
	  val
	end
	
	#获取文件名
	def getWorkbookName()
	  klog val = KSO_SDK::Application.ActiveWorkbook.Name
	  val
	end
	
	#获取工资表名
	def getWorksheetName()
	  klog val = KSO_SDK::Application.ActiveSheet.Name
	  val
	end
	  
	#为单元格设置下来选值
	def setRangeValidation(address, array)
	  #{"type":3,"value":true,"alertStyle":1,"operator":1,"inCellDropdown":true,"formula1":"123,321,abc","formula2":""} 
	  KSO_SDK::Application.ActiveSheet.Range(address).Validation().Delete()
	  KSO_SDK::Application.ActiveSheet.Range(address).Validation().Add(3, 2, 1, array)
    end
	
    def activeSheet(name)
      KSO_SDK::Application.WorkSheets(name).Activate = false
    end
	
	
    def cancelCopyMode()
      KSO_SDK::Application.CutCopyMode = false
    end
	
	
    # 显示alert
    def showAlert(text)
      btnMask = Qt::MessageBox::about(KSO_SDK.getCurrentMainWindow(), '提示', text)
    end
  
  
  
    #获取插件存储文件路径
    def getStorageDir()
      KSO_SDK.getStorageDir(context)
    end
	
	#
    def fileRename(source, target)
      File.rename(source, target)
    end
	
    def fileDelete(file)
      File.delete(file)
    end
	
    def fileCopy(source, target)
	  FileUtils.cp(source, target)
    end
	
	def appid()
		context.appId
	end
	
	def windowScrollTo(row, col)
		KSO_SDK::Application.ActiveWindow.ScrollColumn = col
		KSO_SDK::Application.ActiveWindow.ScrollRow = row
	end
	
    def addEmptyWorkbook()
      KSO_SDK::Application.Workbooks.Add
    end
	
	
	
    #添加Sheet
    def testUndo()
      KSO_SDK::Application.OnUndo("test undo", "testUndoDo")
	  p "set undo"
    end
	
	def testUndoDo()
		p "haha"
	end
	
	
  
  
    def getFileName(line = false)
      if line
        result = __FILE__ + getLine().to_s
      else
        result = __FILE__
      end
      result
    end
	
    def callback(methodName)
      klog methodName
      json = {:params => "content"}.to_json()
      klog json
      callbackToJS(methodName, json)
    end
  
    def test(url)
      command = KSO_SDK::getCurrentMainWindow().commands().command("CT_Home");
      puts command.class
      cmd = KRbTabCommand.new(KSO_SDK::getCurrentMainWindow(), KSO_SDK::getCurrentMainWindow())
      cmd.setDrawText(url)
      KSO_SDK::getCurrentMainWindow().commands().addCommand("CT_MyuHome", cmd)
      "call test"
    end
	
    private
  
    def getLine()
      __LINE__.to_s
    end
  
  end
  
end