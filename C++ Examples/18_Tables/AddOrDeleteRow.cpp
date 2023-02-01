#include "pch.h"
using namespace Spire::Doc;

int main()
{
	wstring input_path = DATAPATH;
	wstring inputFile = input_path + L"TableSample.docx";
	wstring output_path = OUTPUTPATH;
	wstring outputFile = output_path + L"AddOrDeleteRow.docx";

	//Create a document
	Document* document = new Document();
	//Load file
	document->LoadFromFile(inputFile.c_str());
	Section* section = document->GetSections()->GetItem(0);
	Table* table = dynamic_cast<Table*>(section->GetTables()->GetItemInTableCollection(0));

	//Delete the seventh row
	table->GetRows()->RemoveAt(7);

	//Add a row and insert it into specific position
	TableRow* row = new TableRow(document);
	for (int i = 0; i < table->GetRows()->GetItem(0)->GetCells()->GetCount(); i++)
	{
		TableCell* tc = row->AddCell();
		Paragraph* paragraph = tc->AddParagraph();
		paragraph->GetFormat()->SetHorizontalAlignment(HorizontalAlignment::Center);
		paragraph->AppendText(L"Added");
	}
	table->GetRows()->Insert(2, row);
	//Add a row at the end of table
	table->AddRow();

	//Save to file and launch it
	document->SaveToFile(outputFile.c_str(), FileFormat::Docx);
	document->Close();
	delete document;
}
