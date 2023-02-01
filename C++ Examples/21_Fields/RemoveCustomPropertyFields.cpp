#include "pch.h"
using namespace Spire::Doc;

int main() {
	wstring input_path = DATAPATH;
	wstring inputFile = input_path + L"RemoveCustomPropertyFields.docx";
	wstring output_path = OUTPUTPATH;
	wstring outputFile = output_path + L"RemoveCustomPropertyFields.docx";

	//Create Word document.
	Document* document = new Document();

	//Load the file from disk.
	document->LoadFromFile(inputFile.c_str());

	//Get custom document properties object.
	CustomDocumentProperties* cdp = document->GetCustomDocumentProperties();

	//Remove all custom property fields in the document.
	for (int i = 0; i < cdp->GetCount();/* i++*/)
	{
		cdp->Remove(cdp->GetItem(i)->GetName());
	}

	document->SetIsUpdateFields(true);

	//Save to file.
	document->SaveToFile(outputFile.c_str(), FileFormat::Docx2013);
	document->Close();
	delete document;
}
