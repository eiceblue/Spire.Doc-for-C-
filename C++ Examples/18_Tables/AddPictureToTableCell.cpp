#include "pch.h"
using namespace Spire::Doc;

int main()
{
	wstring input_path = DATAPATH;
	wstring inputFile = input_path + L"TableTemplate.docx";
	wstring output_path = OUTPUTPATH;
	wstring outputFile = output_path + L"AddPictureToTableCell.docx";

	//Load the document
	Document* doc = new Document();
	doc->LoadFromFile(inputFile.c_str());

	//Get the first table from the first section of the document
	Table* table1 = dynamic_cast<Table*>(doc->GetSections()->GetItem(0)->GetTables()->GetItemInTableCollection(0));

	//Add a picture to the specified table cell and set picture size
	DocPicture* picture = table1->GetRows()->GetItem(1)->GetCells()->GetItem(2)->GetParagraphs()->GetItem(0)->AppendPicture((input_path + L"Spire.Doc.png").c_str());
	picture->SetWidth(100);
	picture->SetHeight(100);

	//Save and launch document
	doc->SaveToFile(outputFile.c_str(), FileFormat::Docx);
	doc->Close();
	delete doc;
}
