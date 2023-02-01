#include "pch.h"
using namespace Spire::Doc;

int main() {
	wstring input_path = DATAPATH;
	wstring inputFile = input_path + L"Template.docx";
	wstring output_path = OUTPUTPATH;
	wstring outputFile = output_path + L"CheckFileFormat.txt";

	Document* doc = new Document();
	doc->LoadFromFile(inputFile.c_str());
	wstring fileFormat = L"The file format is ";
	//Check the format info
	switch (doc->GetDetectedFormatType())
	{
	case FileFormat::Doc:
		fileFormat += L"Microsoft Word 97-2003 document.";
		break;
	case FileFormat::Dot:
		fileFormat += L"Microsoft Word 97-2003 template.";
		break;
	case FileFormat::Docx:
		fileFormat += L"Office Open XML WordprocessingML Macro-Free Document.";
		break;
	case FileFormat::Docm:
		fileFormat += L"Office Open XML WordprocessingML Macro-Enabled Document.";
		break;
	case FileFormat::Dotx:
		fileFormat += L"Office Open XML WordprocessingML Macro-Free Template.";
		break;
	case FileFormat::Dotm:
		fileFormat += L"Office Open XML WordprocessingML Macro-Enabled Template.";
		break;
	case FileFormat::Rtf:
		fileFormat += L"RTF format.";
		break;
	case FileFormat::WordML:
		fileFormat += L"Microsoft Word 2003 WordprocessingML format.";
		break;
	case FileFormat::Html:
		fileFormat += L"HTML format.";
		break;
	case FileFormat::WordXml:
		fileFormat += L"Microsoft Word xml format for word 2007-2013.";
		break;
	case FileFormat::Odt:
		fileFormat += L"OpenDocument Text.";
		break;
	case FileFormat::Ott:
		fileFormat += L"OpenDocument Text Template.";
		break;
	case FileFormat::DocPre97:
		fileFormat += L"Microsoft Word 6 or Word 95 format.";
		break;
	default:
		fileFormat += L"Unknown format.";
		break;
	}

	//Save to file.
	wofstream out;
	out.open(outputFile);
	out.flush();
	out << fileFormat.c_str();
	out.close();
	doc->Close();
	delete doc;
}
