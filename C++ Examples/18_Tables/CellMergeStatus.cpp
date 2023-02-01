#include "pch.h"
using namespace Spire::Doc;

int main() {
	wstring input_path = DATAPATH;
	wstring inputFile = input_path + L"CellMergeStatus.docx";
	wstring output_path = OUTPUTPATH;
	wstring outputFile = output_path + L"CellMergeStatus.txt";

	Document* doc = new Document();
	doc->LoadFromFile(inputFile.c_str());

	//Get the first section
	Section* section = doc->GetSections()->GetItem(0);

	//Get the first table in the section
	Table* table = dynamic_cast<Table*>(section->GetTables()->GetItemInTableCollection(0));

	wstring* stringBuidler = new wstring();
	for (int i = 0; i < table->GetRows()->GetCount(); i++)
	{
		TableRow* tableRow = table->GetRows()->GetItem(i);
		for (int j = 0; j < tableRow->GetCells()->GetCount(); j++)
		{
			TableCell* tableCell = tableRow->GetCells()->GetItem(j);
			CellMerge verticalMerge = tableCell->GetCellFormat()->GetVerticalMerge();
			short horizontalMerge = tableCell->GetGridSpan();

			if (verticalMerge == CellMerge::None && horizontalMerge == 1)
			{
				stringBuidler->append(L"Row " + to_wstring(i) + L", cell " + to_wstring(j) + L": ");
				stringBuidler->append(L"This cell isn't merged.\r\n");
			}
			else
			{
				stringBuidler->append(L"Row " + to_wstring(i) + L", cell " + to_wstring(j) + L": ");
				stringBuidler->append(L"This cell is merged.\r\n");
			}
		}
		stringBuidler->append(L"\r\n");
	}

	//Save and launch document
	wofstream write(outputFile);
	write << stringBuidler->c_str();
	write.close();
	doc->Close();
	delete doc;
	delete stringBuidler;
}
