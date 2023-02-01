#include "pch.h"
using namespace Spire::Doc;

int main() {
	wstring output_path = OUTPUTPATH;
	wstring outputFile = output_path + L"AddHorizontalLine.docx";

	//Create Word document.
	Document* doc = new Document();
	Section* sec = doc->AddSection();
	Paragraph* para = sec->AddParagraph();
	para->AppendHorizonalLine();

	//Save and launch document
	doc->SaveToFile(outputFile.c_str(), FileFormat::Docx);
	doc->Close();
	delete doc;
}
