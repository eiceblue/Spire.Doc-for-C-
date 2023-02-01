#include "pch.h"
using namespace Spire::Doc;

int main() {
	wstring input_path = DATAPATH;
	wstring inputFile = input_path + L"TableTemplate.docx";
	wstring output_path = OUTPUTPATH;
	wstring outputFile = output_path + L"CloneRow.docx";

	Document* doc = new Document();
	doc->LoadFromFile(inputFile.c_str());

	//Get the first section
	Section* se = doc->GetSections()->GetItem(0);

	//Get the first row of the first table
	TableRow* firstRow = dynamic_cast<Table*>(se->GetTables()->GetItemInTableCollection(0))->GetRows()->GetItem(0);

	//Copy the first row to clone_FirstRow via TableRow.clone()
	TableRow* clone_FirstRow = firstRow->Clone();

	dynamic_cast<Table*>(se->GetTables()->GetItemInTableCollection(0))->GetRows()->Add(clone_FirstRow);
	//Save and launch document
	doc->SaveToFile(outputFile.c_str(), FileFormat::Docx);
	doc->Close();
	delete doc;
}
