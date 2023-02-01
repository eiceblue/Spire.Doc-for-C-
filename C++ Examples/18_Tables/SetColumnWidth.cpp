#include "pch.h"
using namespace Spire::Doc;

int main() {
	wstring input_path = DATAPATH;
	wstring inputFile = input_path + L"TableSample.docx";
	wstring output_path = OUTPUTPATH;
	wstring outputFile = output_path + L"SetColumnWidth.docx";

	//Create a document and load file
	Document* document = new Document();
	document->LoadFromFile(inputFile.c_str());

	Section* section = document->GetSections()->GetItem(0);
	Table* table = dynamic_cast<Table*>(section->GetTables()->GetItemInTableCollection(0));

	//Traverse the first column
	for (int i = 0; i < table->GetRows()->GetCount(); i++)
	{
		//Set the width and type of the cell
		table->GetRows()->GetItem(i)->GetCells()->GetItem(0)->SetCellWidth(200, CellWidthType::Point);
	}

	//Save to file
	document->SaveToFile(outputFile.c_str(), FileFormat::Docx);
	document->Close();
	delete document;
}
