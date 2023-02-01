#include "pch.h"
using namespace Spire::Doc;

int main() {
	wstring input_path = DATAPATH;
	wstring inputFile = input_path + L"Template_Docx_1.docx";
	wstring output_path = OUTPUTPATH;
	wstring outputFile = output_path + L"RemoveSpecificParagraph.docx";

	//Create Word document.
	Document* document = new Document();

	//Load the file from disk.
	document->LoadFromFile(inputFile.c_str());
	//Remove the first paragraph from the first section of the document.
	document->GetSections()->GetItem(0)->GetParagraphs()->RemoveAt(0);

	//Save to file.
	document->SaveToFile(outputFile.c_str(), FileFormat::Docx2013);
	document->Close();
	delete document;
}