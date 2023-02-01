#include "pch.h"
using namespace Spire::Doc;

int main()
{
	//Combine tables
	CombineTables();
	//Split a table
	SplitTable();
}
void CombineTables() {
	wstring input_path = DATAPATH;
	wstring inputFile = input_path + L"CombineAndSplitTables.docx";
	wstring output_path = OUTPUTPATH;
	wstring outputFile = output_path + L"CombineTables.docx";
	
	//Load document from disk
	Document* doc = new Document();
	doc->LoadFromFile(inputFile.c_str());

	//Get the first section
	Section* section = doc->GetSections()->GetItem(0);

	//Get the first and second table
	Table* table1 = dynamic_cast<Table*>(section->GetTables()->GetItemInTableCollection(0));
	Table* table2 = dynamic_cast<Table*>(section->GetTables()->GetItemInTableCollection(1));

	//Add the rows of table2 to table1
	for (int i = 0; i < table2->GetRows()->GetCount(); i++)
	{
		table1->GetRows()->Add(table2->GetRows()->GetItem(i)->Clone());
	}

	//Remove the table2
	section->GetTables()->Remove(table2);

	//Save the Word file
	section->GetDocument()->SaveToFile(outputFile.c_str(), FileFormat::Docx2013);
	doc->Close();
	delete doc;
}

void SplitTable()
{
	wstring input_path = DATAPATH;
	wstring inputFile = input_path + L"CombineAndSplitTables.docx";
	wstring output_path = OUTPUTPATH;
	wstring outputFile = output_path + L"SplitTable.docx";

	//Load document from disk
	Document* doc = new Document();
	doc->LoadFromFile(inputFile.c_str());

	//Get the first section
	Section* section = doc->GetSections()->GetItem(0);

	//Get the first table
	Table* table = dynamic_cast<Table*>(section->GetTables()->GetItemInTableCollection(0));

	//We will split the table at the third row;
	int splitIndex = 2;

	//Create a new table for the split table
	Table* newTable = new Table(section->GetDocument());

	//Add rows to the new table
	for (int i = splitIndex; i < table->GetRows()->GetCount(); i++)
	{
		newTable->GetRows()->Add(table->GetRows()->GetItem(i)->Clone());
	}

	//Remove rows from original table
	for (int i = table->GetRows()->GetCount() - 1; i >= splitIndex; i--)
	{
		table->GetRows()->RemoveAt(i);
	}

	//Add the new table in section
	section->GetTables()->Add(newTable);

	//Save the Word file
	section->GetDocument()->SaveToFile(outputFile.c_str(), FileFormat::Docx2013);
	doc->Close();
	delete doc;
}
