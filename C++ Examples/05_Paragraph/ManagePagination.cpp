#include "pch.h"
using namespace Spire::Doc;

int main() {
	wstring input_path = DATAPATH;
	wstring inputFile = input_path + L"Template_Docx_1.docx";
	wstring output_path = OUTPUTPATH;
	wstring outputFile = output_path + L"ManagePagination.docx";

	//Create Word document.
	Document* document = new Document();

	//Load the file from disk.
	document->LoadFromFile(inputFile.c_str());

	//Get the first section and the paragraph we want to manage the pagination.
	Section* sec = document->GetSections()->GetItem(0);
	Paragraph* para = sec->GetParagraphs()->GetItem(4);

	//Set the pagination format as Format.PageBreakBefore for the checked paragraph.
	para->GetFormat()->SetPageBreakBefore(true);

	//Save the file.
	document->SaveToFile(outputFile.c_str(), FileFormat::Docx2013);
	document->Close();
	delete document;
}
