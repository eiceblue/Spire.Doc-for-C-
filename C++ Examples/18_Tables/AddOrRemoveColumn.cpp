#include "pch.h"
using namespace Spire::Doc;

int main()
{
	wstring input_path = DATAPATH;
	wstring inputFile = input_path + L"Template_N2.docx";
	wstring output_path = OUTPUTPATH;
	wstring outputFile = output_path + L"AddOrRemoveColumn.docx";

	//Load the document from disk.
	Document* doc = new Document();
	doc->LoadFromFile(inputFile.c_str());

	//Access the first section
	Section* section = doc->GetSections()->GetItem(0);

	//Access the first table
	Table* table = dynamic_cast<Table*>(section->GetTables()->GetItemInTableCollection(0));

	//Add a blank column
	int columnIndex1 = 0;
	AddColumn(table, columnIndex1);

	//Remove a column
	int columnIndex2 = 2;
	RemoveColumn(table, columnIndex2);

	//Save the Word file
	doc->SaveToFile(outputFile.c_str(), FileFormat::Docx2013);
	doc->Close();
	delete doc;
}

void AddColumn(Table* table, int columnIndex)
{
	for (int r = 0; r < table->GetRows()->GetCount(); r++)
	{
		TableCell* addCell = new TableCell(table->GetDocument());
		table->GetRows()->GetItem(r)->GetCells()->Insert(columnIndex, addCell);
	}
}

void RemoveColumn(Table* table, int columnIndex)
{
	for (int r = 0; r < table->GetRows()->GetCount(); r++)
	{
		table->GetRows()->GetItem(r)->GetCells()->RemoveAt(columnIndex);
	}
}
