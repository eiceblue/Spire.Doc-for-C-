#include "pch.h"
using namespace Spire::Doc;

int main() {
	wstring output_path = OUTPUTPATH;
	wstring outputFile = output_path + L"InsertRtfStringToDoc.docx";

	//Create Word document.
	Document* document = new Document();

	//Add a new section.
	Section* section = document->AddSection();

	//Add a paragraph to the section.
	Paragraph* para = section->AddParagraph();

	//Declare a String variable to store the Rtf string.
	wstring rtfString = L"{\\rtf1\\ansi\\deff0 {\\fonttbl {\\f0 hakuyoxingshu7000;}}\\f0\\fs28 Hello, World}";

	//Append Rtf string to paragraph.
	para->AppendRTF(rtfString.c_str());

	//Save to file.
	document->SaveToFile(outputFile.c_str(), FileFormat::Docx);
	document->Close();
	delete document;
}