#include "pch.h"
using namespace Spire::Doc;

int main() {
	wstring input_path = DATAPATH;
	wstring inputFile = input_path + L"Template_Docx_1.docx";
	wstring output_path = OUTPUTPATH;
	wstring outputFile = output_path + L"RemoveAllParagraphs.docx";

	//Create Word document.
	Document* document = new Document();

	//Load the file from disk.
	document->LoadFromFile(inputFile.c_str());

	//Remove paragraphs from every section in the document
	for (int i = 0; i < document->GetSections()->GetCount(); i++)
	{
		Section* section = document->GetSections()->GetItem(i);
		section->GetParagraphs()->Clear();
	}

	//Save to file.
	document->SaveToFile(outputFile.c_str(), FileFormat::Docx2013);
	document->Close();
	delete document;
}