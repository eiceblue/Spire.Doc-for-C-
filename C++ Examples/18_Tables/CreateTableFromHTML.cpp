#include "pch.h"
using namespace Spire::Doc;

int main() {
	wstring output_path = OUTPUTPATH;
	wstring outputFile = output_path + L"CreateTableFromHTML.docx";

	//HTML string
	wstring HTML = L"<table border='2px'><tr><td>Row 1, Cell 1</td><td>Row 1, Cell 2</td></tr><tr><td>Row 2, Cell 2</td><td>Row 2, Cell 2</td></tr></table>";

	//Create a Word document
	Document* document = new Document();

	//Add a section
	Section* section = document->AddSection();

	//Add a paragraph and append html string
	section->AddParagraph()->AppendHTML(HTML.c_str());

	//Save to Word document
	document->SaveToFile(outputFile.c_str(), FileFormat::Docx2013);
	document->Close();
	delete document;
}
