#include "pch.h"
using namespace Spire::Doc;

int main() {
	wstring input_path = DATAPATH;
	wstring inputFile = input_path + L"TableSample.docx";
	wstring output_path = OUTPUTPATH;
	wstring outputFile = output_path + L"MergeAndSplitTableCell.docx";

	//Create a document and load file from disk
	Document* document = new Document();
	document->LoadFromFile(inputFile.c_str());
	Section* section = document->GetSections()->GetItem(0);
	Table* table = dynamic_cast<Table*>(section->GetTables()->GetItemInTableCollection(0));
	//The method shows how to merge cell horizontally
	table->ApplyHorizontalMerge(6, 2, 3);
	//The method shows how to merge cell vertically
	table->ApplyVerticalMerge(2, 4, 5);
	//The method shows how to split the cell
	table->GetRows()->GetItem(8)->GetCells()->GetItem(3)->SplitCell(2, 2);
	//Save to file and launch it
	document->SaveToFile(outputFile.c_str(), FileFormat::Docx);
	document->Close();
	delete document;
}
