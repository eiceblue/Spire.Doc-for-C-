#include "pch.h"
using namespace Spire::Doc;

int main() {
	wstring input_path = DATAPATH;
	wstring inputFile = input_path + L"ExtractTextFromTextBoxes.docx";
	wstring output_path = OUTPUTPATH;
	wstring outputFile = output_path + L"ExtractTextFromTextBoxes.txt";

	//Create a Word document.
	Document* document = new Document();

	//Load the file from disk.
	document->LoadFromFile(inputFile.c_str());

	//Verify whether the document contains a textbox or not.
	if (document->GetTextBoxes()->GetCount() > 0)
	{
		wofstream sw(outputFile);
		//Traverse the document.
		for (int i = 0; i < document->GetSections()->GetCount(); i++)
		{
			Section* section = document->GetSections()->GetItem(i);
			for (int j = 0; j < section->GetParagraphs()->GetCount(); j++)
			{
				Paragraph* p = section->GetParagraphs()->GetItem(j);
				for (int k = 0; k < p->GetChildObjects()->GetCount(); k++)
				{
					DocumentObject* obj = p->GetChildObjects()->GetItem(k);
					if (obj->GetDocumentObjectType() == DocumentObjectType::TextBox)
					{
						TextBox* textbox = dynamic_cast<TextBox*>(obj);
						for (int l = 0; l < textbox->GetChildObjects()->GetCount(); l++)
						{
							DocumentObject* objt = textbox->GetChildObjects()->GetItem(l);
							//Extract text from paragraph in TextBox.
							if (objt->GetDocumentObjectType() == DocumentObjectType::Paragraph)
							{
								sw << ((dynamic_cast<Paragraph*>(objt))->GetText());
							}

							//Extract text from Table in TextBox.
							if (objt->GetDocumentObjectType() == DocumentObjectType::Table)
							{
								Table* table = dynamic_cast<Table*>(objt);
								ExtractTextFromTables(table, sw);
							}
						}
					}
				}
			}
		}
	}
	document->Close();
	delete document;
}

void ExtractTextFromTables(Table* table, wofstream& sw)
{
	for (int i = 0; i < table->GetRows()->GetCount(); i++)
	{
		TableRow* row = table->GetRows()->GetItem(i);
		for (int j = 0; j < row->GetCells()->GetCount(); j++)
		{
			TableCell* cell = row->GetCells()->GetItem(j);
			for (int k = 0; k < cell->GetParagraphs()->GetCount(); k++)
			{
				Paragraph* paragraph = cell->GetParagraphs()->GetItem(k);
				sw << (paragraph->GetText());
			}
		}
	}
}