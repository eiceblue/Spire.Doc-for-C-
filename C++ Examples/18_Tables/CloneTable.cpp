#include "pch.h"
using namespace Spire::Doc;

int main() {
	wstring input_path = DATAPATH;
	wstring inputFile = input_path + L"TableTemplate.docx";
	wstring output_path = OUTPUTPATH;
	wstring outputFile = output_path + L"CloneTable.docx";

	//Load the document
	Document* doc = new Document();
	doc->LoadFromFile(inputFile.c_str());

	//Get the first section
	Section* se = doc->GetSections()->GetItem(0);

	//Get the first table
	Table* original_Table = dynamic_cast<Table*>(se->GetTables()->GetItemInTableCollection(0));

	//Copy the existing table to copied_Table via Table.clone()
	Table* copied_Table = original_Table->Clone();
	vector<wstring> st = { L"Spire.Presentation for .Net", L"A professional PowerPoint® compatible library that enables developers to create, read, write, modify, convert and Print PowerPoint documents on any C++ framework, C++ Core platform." };
	//Get the last row of table
	TableRow* lastRow = copied_Table->GetRows()->GetItem(copied_Table->GetRows()->GetCount() - 1);
	//Change last row data
	for (int i = 0; i < lastRow->GetCells()->GetCount() - 1; i++)
	{
		lastRow->GetCells()->GetItem(i)->GetParagraphs()->GetItem(0)->SetText(st[i].c_str());
	}
	//Add copied_Table in section
	se->GetTables()->Add(copied_Table);

	//Save and launch document
	doc->SaveToFile(outputFile.c_str(), FileFormat::Docx);
	doc->Close();
	delete doc;
}
