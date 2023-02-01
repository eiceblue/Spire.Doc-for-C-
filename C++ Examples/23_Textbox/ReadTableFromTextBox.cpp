#include "pch.h"
using namespace Spire::Doc;

int main() {
	wstring input_path = DATAPATH;
	wstring inputFile = input_path + L"TextBoxTable.docx";
	wstring output_path = OUTPUTPATH;
	wstring outputFile = output_path + L"ReadTableFromTextBox.txt";

	//Load the document
	Document* doc = new Document();
	doc->LoadFromFile(inputFile.c_str());

	//Get the first textbox
	TextBox* textbox = doc->GetTextBoxes()->GetItem(0);

	//Get the first table in the textbox
	Table* table = dynamic_cast<Table*>(textbox->GetBody()->GetTables()->GetItemInTableCollection(0));

	wstring* string_builder = new wstring();

	//Loop through the paragraphs of the table cells and extract them to a .txt file
	for (int i = 0; i < table->GetRows()->GetCount(); i++)
	{
		TableRow* row = table->GetRows()->GetItem(i);
		for (int j = 0; j < row->GetCells()->GetCount(); j++)
		{
			TableCell* cell = row->GetCells()->GetItem(j);
			for (int k = 0; k < cell->GetParagraphs()->GetCount(); k++)
			{
				Paragraph* paragraph = cell->GetParagraphs()->GetItem(k);
				string_builder->append(paragraph->GetText());
				string_builder->append(L"\t");
			}
		}
		string_builder->append(L"\n");
	}

	//Save to TXT file and launch it
	wofstream foo(outputFile);
	foo << string_builder->c_str();
	foo.close();
	doc->Close();
	delete doc;
}
