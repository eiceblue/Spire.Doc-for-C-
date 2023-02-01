#include "pch.h"
using namespace Spire::Doc;

int main() {
	wstring input_path = DATAPATH;
	wstring inputFile = input_path + L"Template_Docx_5.docx";
	wstring output_path = OUTPUTPATH;
	wstring outputFile = output_path + L"PreventPageBreakInTable.docx";

	//Create Word document.
	Document* document = new Document();

	//Load the file from disk.
	document->LoadFromFile(inputFile.c_str());

	//Get the table from Word document.
	Table* table = dynamic_cast<Table*>(document->GetSections()->GetItem(0)->GetTables()->GetItemInTableCollection(0));

	//Change the paragraph setting to keep them together.
	for (int i = 0; i < table->GetRows()->GetCount(); i++)
	{
		TableRow* row = table->GetRows()->GetItem(i);
		for (int j = 0; j < row->GetCells()->GetCount(); j++)
		{
			TableCell* cell = row->GetCells()->GetItem(j);
			for (int k = 0; k < cell->GetParagraphs()->GetCount(); k++)
			{
				Paragraph* p = cell->GetParagraphs()->GetItem(k);
				p->GetFormat()->SetKeepFollow(true);
			}
		}
	}

	//Save to file.
	document->SaveToFile(outputFile.c_str(), FileFormat::Docx2013);
	document->Close();
	delete document;
}
