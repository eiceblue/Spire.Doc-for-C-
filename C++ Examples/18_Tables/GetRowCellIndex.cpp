#include "pch.h"
using namespace Spire::Doc;

int main() {
	wstring input_path = DATAPATH;
	wstring inputFile = input_path + L"ReplaceTextInTable.docx";
	wstring output_path = OUTPUTPATH;
	wstring outputFile = output_path + L"GetRowCellIndex.txt";

	//Load Word from disk
	Document* doc = new Document();
	doc->LoadFromFile(inputFile.c_str());

	//Get the first section
	Section* section = doc->GetSections()->GetItem(0);

	//Get the first table in the section
	Table* table = dynamic_cast<Table*>(section->GetTables()->GetItemInTableCollection(0));

	wstring* content = new wstring();

	//Get table collections
	TableCollection* collections = section->GetTables();

	//Get the table index
	int tableIndex = collections->IndexOf(table);

	//Get the index of the last table row
	TableRow* row = table->GetLastRow();
	int rowIndex = row->GetRowIndex();

	//Get the index of the last table cell
	TableCell* cell = dynamic_cast<TableCell*>(row->GetLastChild());
	int cellIndex = cell->GetCellIndex();

	//Append these information into content
	content->append(L"Table index is " + to_wstring(tableIndex));
	content->append(L"\nRow index is " + to_wstring(rowIndex));
	content->append(L"\nCell index is " + to_wstring(cellIndex));

	//Save to txt file
	wofstream write(outputFile);
	write << content->c_str();
	write.close();
	doc->Close();
	delete doc;
	delete content;

}
