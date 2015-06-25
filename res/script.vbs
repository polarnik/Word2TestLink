Dim objWord, i, j, iStop, objXmlDoc, p, Doc, Root, Node, NodeDef, startDef, endDef, defRande, Text, Attr, objXslDoc, str, FSO, level, testDef, isTestCase, testRange, testRangeEnd, testRangeStart, parIndex, parNode, tableNode, testStepsTable, testStepRow, testStepText, stepResultText, testStepRowIndex, tmp, tmpFolderPath, extractPictureCount

Function SaveTextNode(textRange, defNode)
	Dim textNode, brNode
	str = textRange.Text
	badcode = 0
	while (badcode <= 7)
		str = Replace(str, Chr(badcode), "")
		badcode = badcode + 1
	wend
	strClean = Replace(str, vbCR, " ")
	Set textNode = objXmlDoc.createTextNode(strClean)
	defNode.appendChild textNode
End Function

Function SaveImageNode(imageRange, defNode)
	Dim defParagraph, InlineShape, InlineShapeRange, imageXML, binDataNodes, imageTag, DM, EL, binaryStream, imgHtmlTag, srcHtmlAttr, heightHtmlAttr, widthHtmlAttr, aHtmlTag, hrefHtmlAttr, targetHtmlAttr
	if ((document.getElementById("processPicture").checked)) then
		if(imageRange.InlineShapes(1).Type = 1) then
			Set InlineShape = imageRange.InlineShapes(1)
			Set InlineShapeRange = InlineShape.Range 'Doc.Range(InlineShape.Range.Start - 1, InlineShape.Range.End + 1)
			Set imageXML = CreateObject("msxml.DomDocument")
			imageXML.LoadXML InlineShapeRange.XML
			Set binDataNodes = imageXML.getElementsByTagName("w:binData")
			if(binDataNodes.length > 0) then
				'В этом теге лежит png-картинка в формате Base64
				Set imageTag = binDataNodes.item(binDataNodes.length - 1)
				'Раскодируем картинку из Base64 в набор байтов
				'алгоритм обработки взят отсюда: https://ghads.wordpress.com/2008/10/17/vbscript-readwrite-binary-encodedecode-base64/
				Set DM = CreateObject("Microsoft.XMLDOM")
				Set EL = DM.createElement("tmp")
				EL.DataType = "bin.base64"
				EL.Text = imageTag.Text
				decodeBase64 = EL.NodeTypedValue
				'Сохраним картинку в файл
				Const TypeBinary = 1
				Const ForReading = 1, ForWriting = 2, ForAppending = 8
				Set binaryStream = CreateObject("ADODB.Stream") 
				binaryStream.Type = TypeBinary
				binaryStream.Open
				binaryStream.Write decodeBase64
				'Сформируем хеш от содержимого файла, значение хеша будет уникальным именем файла
				'алгоритм получения хеша основан на http://www.daleanderson.ca/code/wsh/md5test.vbs.txt
				Const CAPICOM_HASH_ALGORITHM_MD5 =       3
				Dim objHashedData
				Set objHashedData = CreateObject("CAPICOM.HashedData")
				objHashedData.Algorithm = CAPICOM_HASH_ALGORITHM_MD5
				objHashedData.Hash decodeBase64
				'alert(objHashedData.Value)
				'Указываем хеш как имя файла
				file = tmpFolderPath & "\" & objHashedData.Value & ".png"
				
				binaryStream.SaveToFile file, ForWriting
				extractPictureCount = extractPictureCount + 1
				'alert(decodeBase64)
				Set imgHtmlTag = objXmlDoc.CreateNode(1, "img", "")
				Set srcHtmlAttr = objXmlDoc.CreateNode(2, "src", "")
				href = "\userfiles\word2TestLink\" & objHashedData.Value & ".png"
				srcHtmlAttr.NodeValue = href
				imgHtmlTag.attributes.setNamedItem srcHtmlAttr
				
				'Set heightHtmlAttr = objXmlDoc.CreateNode(2, "height", "")
				'heightHtmlAttr.nodeValue = imageRange.InlineShapes(1).Height
				'imgHtmlTag.attributes.setNamedItem heightHtmlAttr
				
				Set widthHtmlAttr = objXmlDoc.CreateNode(2, "width", "")
				widthHtmlAttr.nodeValue = imageRange.InlineShapes(1).width
				imgHtmlTag.attributes.setNamedItem widthHtmlAttr
				
				Set aHtmlTag = objXmlDoc.CreateNode(1, "a", "")
				Set hrefHtmlAttr = objXmlDoc.CreateNode(2, "href", "")
				hrefHtmlAttr.NodeValue = href
				aHtmlTag.attributes.setNamedItem hrefHtmlAttr
				Set targetHtmlAttr = objXmlDoc.CreateNode(2, "target", "")
				targetHtmlAttr.NodeValue = "_blank"
				aHtmlTag.attributes.setNamedItem targetHtmlAttr

				aHtmlTag.appendChild imgHtmlTag
				defNode.appendChild aHtmlTag
			end if

			'imageXML.Save "E:\Data\DEVEL\Word2TestLink.1.0.4\img\img.xml"
			'alert(imageXML.XML)
		end if
	end if
End Function

'Разбор текста одного параграфа (описания теста или группы тестов, текст шага или ожидаемого результата)
'Разбор форматирования в этом параграфе, и формирование html-кода (секции CDATA), который можно вставить в описание или в шаг.
Function GetCDATADefenition(defRange, defNode)
	'Надо предсмотреть поддержку:
	'абзацев
	'жирного текста
	'курсивного текста
	'подчёркнутого текста
	'зачеркнутого текста
	'картинок
	Dim defParagraph, InlineShape, InlineShapeRange, imageXML, binDataNodes, imageTag, DM, EL, binaryStream
	Set defParagraph = objXmlDoc.CreateNode(1, "p", "")

	Dim currRange
	curStart = defRange.Start
	curEnd = curStart + 1
	Set currRange = Doc.Range(curStart, curStart)
	while curEnd <= defRange.End
		Set currRange = Doc.Range(curStart, curEnd)
		if(currRange.InlineShapes.count > 0) then
			if(currRange.InlineShapes(1).Type = 1) then
				SaveTextNode currRange, defParagraph
				SaveImageNode currRange, defParagraph
				curStart = curEnd
			end if
		end if
		curEnd = curEnd + 1
	WEnd
	SaveTextNode currRange, defParagraph

	defNode.appendChild defParagraph
	
	
End Function


Function StartConvert()
Dim TestRulesLevel
document.getElementById("testLink").value = "Заполните поле ""Путь к исходному файлу doc, docx, ...""."
If ((document.getElementById("fileName").value <> "")) Then
	Set FSO = CreateObject("Scripting.FileSystemObject")
	tmpFolderPath = FSO.GetSpecialFolder(2).Path & "\" & FSO.GetTempName()
	FSO.CreateFolder tmpFolderPath
	extractPictureCount = 0
		
	document.getElementById("testLink").value = "Преобразование Word 2 TestLink начато."
	Set objWord = CreateObject("Word.Application")
	if ((document.getElementById("visibleForWord").checked)) then
		objWord.Visible = True
	End if
	set Doc = objWord.Documents.Open (document.getElementById("fileName").value, , , , , , , , , , false, , , , true)
	Set p = Doc.Paragraphs
	i = 1
	iStop = p.Count
	Set objXmlDoc = CreateObject("msxml.DomDocument")
	objXmlDoc.loadXML "<root></root>"
	Set Root = objXmlDoc.firstChild
	Dim RootTitle
	Set RootTitle = objXmlDoc.CreateNode(2, "title", "")
	RootTitle.nodeValue = Replace(p(1).Range.Text, vbCR, "")
	Root.attributes.setNamedItem RootTitle
	pokaNeNaidenZagolovok = true
	While (i <= iStop) and pokaNeNaidenZagolovok
		if ((InStr(Trim(UCase(p(i).Range.Text)),  "ТЕСТОВЫЕ УСЛОВИЯ")>0) and (p(i).Format.OutlineLevel >=1) and (p(i).Format.OutlineLevel <=3)) then
			'alert(p(i).Range.Text)
			Set Attr = objXmlDoc.CreateNode(2, "indent", "")
			Attr.nodeValue = p(i).Format.OutlineLevel
			Root.attributes.setNamedItem Attr
			pokaNeNaidenZagolovok = false
		else
		end if
		i = i + 1
	WEnd
	'alert("i:" & i)
	TestRulesLevel = -1
	if(i <= iStop) and (pokaNeNaidenZagolovok = false) then
		TestRulesLevel = p(i-1).Format.OutlineLevel
		'alert("TestRulesLevel:"&TestRulesLevel)
	end if
	While (i <= iStop)
		if(p(i).Format.OutlineLevel > TestRulesLevel) then
			Set Node = objXmlDoc.CreateNode(1, "testnode", "")
			p(i).Range.Select
			'Node.text = p(i).Range.Text
			'level = p(i).Format.OutlineLevel
			if ((p(i).Format.OutlineLevel >= 1) and (p(i).Format.OutlineLevel <= 9)) Then
				Set Attr = objXmlDoc.CreateNode(2, "indent", "")
				Attr.nodeValue = p(i).Format.OutlineLevel
				Node.attributes.setNamedItem Attr
				Set Attr = objXmlDoc.CreateNode(2, "title", "")
				str = p(i).Range.Text
				badcode = 0
				while (badcode <= 7)
					str = Replace(str, Chr(badcode), "")
					badcode = badcode + 1
				wend
				str = Replace(str, vbCR, " ")
				
				Attr.nodeValue = str
				Node.attributes.setNamedItem Attr
				Root.appendChild Node
				
				'Тест или группа тестов в иерархию добавлены, теперь надо достать описание этих тестов
				testDef = true
				isTestCase = -1
				startDef = -1
				endDef = -1
				j = i + 1
				While (j <= iStop) and testDef
					if(p(j).Format.OutlineLevel = 10) Then
						if (startDef = -1) Then
							startDef = p(j).Range.Start
						End If
						endDef = p(j).Range.End
						p(j).Range.Select
						j = j + 1
					Else
						testDef = false
						if (p(j).Format.OutlineLevel < 10) Then
							If (p(i).Format.OutlineLevel < p(j).Format.OutlineLevel) Then
								isTestCase = 0 ' это не TestCase, это TestSuite
							Else
								isTestCase = 1 ' это TestCase - так как следующий элемент старше, чем текущий или равен по уровню
							End If
						End if
					End If
				Wend
				if (startDef > -1) and (endDef > -1) Then
					Set defRande = Doc.Range(startDef, endDef)
					ParseTestDefinition defRande, isTestCase, objXmlDoc, Node 
				
					Node.appendChild NodeDef
				End If
				i = j
			Else
				'Если уровень не задан, то это не тест и не группа тестов, а описание предыдущего теста.
				i = i + 1
			End If
		else
			i = iStop + 1
		End If
	Wend
	'objXmlDoc.Save "%temp%/original.xml"
	'objXmlDoc.Save "E:\Data\DEVEL\Word2TestLink.1.0.4\original.xml"
	Set objXslDoc = CreateObject("msxml.DomDocument")
	objXslDoc.Load "res/style.1.0.5.xsl"
	objXmlDoc.loadXML objXmlDoc.transformNode(objXslDoc)
	testLink_fileName = tmpFolderPath & "\Test_Suite.xml"
	objXmlDoc.Save testLink_fileName
	document.getElementById("testLink").value = testLink_fileName
	
	if (extractPictureCount = 0) then
		document.getElementById("processPictureInfo").innerHTML = "В из документа не были извлечены изображения (количество изображений = 0)."
	else
		document.getElementById("processPictureInfo").innerHTML = "<div>Количество извлечённых изображений: " & extractPictureCount & ".</div><div>Теперь надо скопировать все картинки из папки</div><div><input type=""text"" readOnly=""true"" title=""Сюда изображения были извлечены"" style=""width: 100%"" value=""" & tmpFolderPath & """ /></div><div>в папку</div><div><input type=""text"" readOnly=""true"" title=""Сюда изображения были извлечены"" style=""width: 100%"" value=""\\{{{SERVER}}}\C$\inetpub\testlink\userfiles\word2TestLink\"" /></div><div>(скрипт не может это сделать автоматически из-за ограчений безопасности)</div>"
	End if
	Doc.Close false
	Set Node = Nothing
	Set Attr = Nothing
	Set Root = Nothing
	Set objXslDoc = Nothing
	Set objXmlDoc = Nothing
	objWord.Visible = False
	objWord.Quit false
	Set objWord = Nothing
End If
End Function

'Функция получает на вход объект Range (кусок текста из документа Word)
'Этот кусок теста является 
Function ParseTestDefinition(inputRange, isTestCaseRange, objXmlDoc, Node)
	'Обработка для TestCase - описание, предусловия, шаги с результатами и трудоёмкость
	if isTestCaseRange = 1 Then
		Set NodeDef = objXmlDoc.CreateNode(1, "def", "")
		Node.appendChild NodeDef
		if (inputRange.Tables.Count > 0) Then
			'получаем текст до таблицы
			testRangeStart = inputRange.Start
			testRangeEnd = testRangeStart
			Set testRange = Doc.Range(testRangeStart, testRangeEnd)
			While (testRange.Tables.Count = 0) and (testRangeEnd <= inputRange.End)
				testRangeEnd = testRangeEnd + 1
				Set testRange = Doc.Range(testRangeStart, testRangeEnd)
			Wend
			testRangeEnd = testRangeEnd - 1
			'если есть текст до таблицы, то начальная позиция будет меньше конечной
			if (testRangeStart < testRangeEnd) then
				Set testRange = Doc.Range(testRangeStart, testRangeEnd)
				'обрабатываем полученный текст
				parIndex = 1
				While parIndex <= testRange.Paragraphs.Count 
				
					Set parNode = objXmlDoc.CreateNode(1, "par", "")
					parNode.Text = Trim(Replace(testRange.Paragraphs(parIndex).Range.Text, vbCR, " "))
					NodeDef.appendChild parNode
					Set parNode = Nothing
					parIndex = parIndex + 1
				Wend
			end if
			'Теперь обработка таблицы
			Set testStepsTable = inputRange.Tables(1)
			if ((testStepsTable.Columns.Count >= 2) and (testStepsTable.Rows.Count >= 2)) then
				Set tableNode = objXmlDoc.CreateNode(1, "testSteps", "")
				testStepRowIndex = 2
				while (testStepRowIndex <= testStepsTable.Rows.Count)
					Set testStepRow = objXmlDoc.CreateNode(1, "step", "")
					
					'Set testStepText = objXmlDoc.CreateNode(2, "text", "")
					'testStepText.nodeValue = Trim(Replace(Replace(testStepsTable.Cell(testStepRowIndex,1).Range.text, vbCR, " "), Chr(7), ""))
					'testStepRow.attributes.setNamedItem testStepText
					'DEBUG
					Set testStepText = objXmlDoc.CreateNode(1, "text", "")
					GetCDATADefenition testStepsTable.Cell(testStepRowIndex,1).Range, testStepText
					testStepRow.appendChild testStepText
					
					'Set stepResultText = objXmlDoc.CreateNode(2, "result", "")
					'stepResultText.nodeValue = Trim(Replace(Replace(testStepsTable.Cell(testStepRowIndex,2).Range.text, vbCR, " "), Chr(7), ""))
					'testStepRow.attributes.setNamedItem stepResultText

					Set stepResultText = objXmlDoc.CreateNode(1, "result", "")
					GetCDATADefenition testStepsTable.Cell(testStepRowIndex,2).Range, stepResultText
					testStepRow.appendChild stepResultText
					
					tableNode.appendChild testStepRow
					testStepRowIndex = testStepRowIndex + 1
				wend
				Node.appendChild tableNode
			end if
		else
			'NodeDef.text = "TestCase таблиц нет"
			'У теста есть только текстовое описание
			'пробежимся по параграфам и все содержимок добавим в описание
			Dim testDefPar
			Set testDefPar = inputRange.Paragraphs
			testDefParCount  = testDefPar.Count
			testDefParIndex = 1
			while testDefParIndex <= testDefParCount
				GetCDATADefenition testDefPar(testDefParIndex).Range, NodeDef
				testDefParIndex = testDefParIndex + 1
			WEnd
		End if
		
		'Node.appendChild NodeDef
	else
		if isTestCaseRange = 0 Then
			Set NodeDef =  objXmlDoc.CreateNode(1, "def", "")
			Dim testSuiteDefPar
			Set testSuiteDefPar = inputRange.Paragraphs
			testDefParCount  = testSuiteDefPar.Count
			testDefParIndex = 1
			while testDefParIndex <= testDefParCount
				GetCDATADefenition testSuiteDefPar(testDefParIndex).Range, NodeDef
				testDefParIndex = testDefParIndex + 1
			WEnd
			Node.appendChild NodeDef
		else
			Set NodeDef =  objXmlDoc.CreateNode(1, "def", "")
			Node.appendChild NodeDef
		End if
	End if
End Function

Function ParseAttribute (inputText, pNode)
	if((InStr(inputText, "{")>0)and(InStr(inputText, "}")>0)and(InStr(inputText, "}") > InStr(inputText, "{"))) then
		Set Attr = objXmlDoc.CreateNode(2, "info", "")
		Attr.nodeValue = Right(Left(inputText, InStr(inputText, "}")-1), InStr(inputText, "}")-InStr(inputText, "{")-1)
		pNode.attributes.setNamedItem Attr
	End if

	if((InStr(inputText, "[")>0)and(InStr(inputText, "]")>0)and(InStr(inputText, "]") > InStr(inputText, "["))) then
		Set Attr = objXmlDoc.CreateNode(2, "result", "")
		Attr.nodeValue = Right(Left(inputText, InStr(inputText, "]")-1), InStr(inputText, "]")-InStr(inputText, "[")-1)
		pNode.attributes.setNamedItem Attr
	End if

	if((InStr(inputText, "(*")>0)and(InStr(inputText, "*)")>0)and(InStr(inputText, "*)") > InStr(inputText, "(*"))) then
		Set Attr = objXmlDoc.CreateNode(2, "action", "")
		Attr.nodeValue = Right(Left(inputText, InStr(inputText, "*)")-1), InStr(inputText, "*)")-InStr(inputText, "(*")-2)
		pNode.attributes.setNamedItem Attr
	End if

	Set Attr = objXmlDoc.CreateNode(2, "text", "")
	'Set str = Nothing
	if (InStr(inputText, "{")>0) then
		inputText = Left(inputText, InStr(inputText, "{")-1)
	end if
	if (InStr(inputText, "[")>0) then
		inputText = "" & Left(inputText, InStr(inputText, "[")-1)
	end if
	if (InStr(inputText, "(*")>0) then
		inputText = Left(inputText, InStr(inputText, "(*")-2)
	end if
	Attr.nodeValue =inputText
	pNode.attributes.setNamedItem Attr
	
End Function
