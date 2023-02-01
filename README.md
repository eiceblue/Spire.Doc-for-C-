## Spire.Doc for C++ - A C++ Library for Processing Word Documents

[![Foo](https://i.imgur.com/T2ogReo.png)](https://www.e-iceblue.com/Introduce/doc-for-CPP.html)

[Product Page](https://www.e-iceblue.com/Introduce/doc-for-CPP.html) |  [Forum](https://www.e-iceblue.com/forum/spire-doc-f6.html) | [Customized Demo](https://www.e-iceblue.com/Misc/customized-demo.html) | [Temporary License](https://www.e-iceblue.com/TemLicense.html)

[Spire.Doc for C++](https://www.e-iceblue.com/Introduce/doc-for-CPP.html) is a professional Word C++ library specifically designed for developers to create, read, write, convert, merge, split, and compare Word documents on any C++ platforms with fast and high-quality performance.

As an independent Word C++ API, Spire.Doc for C++ doesn't need Microsoft Word to be installed on neither the development nor target systems. However, it can incorporate Microsoft Word document creation capabilities into any developers' C++ applications.

### 100% Standalone C++ API

Spire.Doc for C++ is a totally independent C++ Word class library which doesn't require Microsoft Office installed on system.

### Richest Word Document Features Support

A common use of Spire.Doc for C++ is to create Word document dynamically from scratch. Almost all Word document elements are supported, including pages, sections, headers, footers, digital signatures, footnotes, paragraphs, lists, tables, text, fields, hyperlinks, bookmarks, comments, images, style, background settings, document settings and protection. Furthermore, drawing objects including shapes, textboxes, images, OLE objects, Latex Math Symbols, MathML Code and controls are supported as well.

### Convert File Documents with High Quality

- Convert Word Doc/Docx to XML, RTF, EMF, TXT, XPS, EPUB, HTML, SVG, ODT
- Convert XML, RTF, EMF, TXT, XPS, EPUB, HTML, SVG, ODT to Word Doc/Docx
- Convert Word Doc/Docx to PDF 
- Convert HTML to Image
- Save Word Doc/Docx to stream
- Save Word Doc/Docx as web response

### Examples

### Create If Field in C++

```c++
#include "pch.h"
using namespace Spire::Doc;

int main() {
	wstring output_path = OUTPUTPATH;
	wstring outputFile = output_path + L"CreateIFField.docx";

	//Create Word document.
	Document* document = new Document();

	//Add a new section.
	Section* section = document->AddSection();

	//Add a new paragraph.
	Paragraph* paragraph = section->AddParagraph();

	// Define a method of creating an IF Field.
	CreateIfField(document, paragraph);

	//Define merged data.
	vector<LPCWSTR_S> fieldName = { L"Count" };
	vector<LPCWSTR_S> fieldValue = { L"2" };

	//Merge data into the IF Field.
	document->GetMailMerge()->Execute(fieldName, fieldValue);

	//Update all fields in the document.
	document->SetIsUpdateFields(true);

	//Save to file.
	document->SaveToFile(outputFile.c_str(), FileFormat::Docx2013);
	document->Close();
	delete document;

}

void CreateIfField(Document* document, Paragraph* paragraph)
{
	IfField* ifField = new IfField(document);
	ifField->SetType(FieldType::FieldIf);
	ifField->SetCode(L"IF ");

	paragraph->GetItems()->Add(ifField);
	paragraph->AppendField(L"Count", FieldType::FieldMergeField);
	paragraph->AppendText(L" > ");
	paragraph->AppendText(L"\"100\" ");
	paragraph->AppendText(L"\"Thanks\" ");
	paragraph->AppendText(L"\"The minimum order is 100 units\"");

	ParagraphBase* end = document->CreateParagraphItem(ParagraphItemType::FieldMark);
	FieldMark* fm = dynamic_cast<FieldMark*>(end);
	fm->SetType(FieldMarkType::FieldEnd);
	paragraph->GetItems()->Add(end);
	ifField->SetEnd(dynamic_cast<FieldMark*>(end));
}
```

### Convert Word to PDF in C++

```c++
#include "pch.h"
using namespace Spire::Doc;

int main() {
	wstring input_path = DATAPATH;
	wstring inputFile = input_path + L"ConvertedTemplate.docx";
	wstring output_path = OUTPUTPATH;
	wstring outputFile = output_path + L"ToPDF.pdf";

	//Create word document
	Document* document = new Document();
	document->LoadFromFile(inputFile.c_str());

	//Save the document to a PDF file.
	document->SaveToFile(outputFile.c_str(), FileFormat::PDF);
	document->Close();
	delete document;
}
```

### Convert Word to Images in C++

```c++
#include "pch.h"
using namespace Spire::Doc;


int main() {
	wstring input_path = DATAPATH;
	wstring inputFile = input_path + L"ConvertedTemplate.docx";
	wstring output_path = OUTPUTPATH;
	wstring outputFile = output_path + L"ToImage.png";

	//Create word document
	Document* document = new Document();
	document->LoadFromFile(inputFile.c_str());

	//Save doc file.
	Stream* imageStream = document->SaveToImages(0, ImageFormat::GetPng());
	imageStream->Save(outputFile.c_str());
	document->Close();
	delete document;
	imageStream->Dispose();
}

```

### Encrypt Word Document in C++

```c++
#include "pch.h"
using namespace Spire::Doc;

int main() {
	wstring input_path = DATAPATH;
	wstring inputFile = input_path + L"Template.docx";
	wstring output_path = OUTPUTPATH;
	wstring outputFile = output_path + L"Encrypt.docx";

	//Create word document
	Document* document = new Document();

	//Load Word document.
	document->LoadFromFile(inputFile.c_str());

	//encrypt document with password specified by textBox1
	document->Encrypt(L"E-iceblue");

	//Save as docx file.
	document->SaveToFile(outputFile.c_str(), FileFormat::Docx);
	document->Close();
	delete document;
}

```

[Product Page](https://www.e-iceblue.com/Introduce/doc-for-CPP.html)  |  [Forum](https://www.e-iceblue.com/forum/spire-doc-f6.html) | [Temporary License](https://www.e-iceblue.com/TemLicense.html) | [Customized Demo](https://www.e-iceblue.com/Misc/customized-demo.html)
