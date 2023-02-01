#include "pch.h"
using namespace Spire::Doc;

int main() {
	wstring input_path = DATAPATH;
	wstring inputFile = input_path + L"Template_Docx_1.docx";
	wstring output_path = OUTPUTPATH;
	wstring outputFile = output_path + L"HideParagraph.docx";

	//Create Word document.
	Document* document = new Document();

	//Load the file from disk.
	document->LoadFromFile(inputFile.c_str());

	//Get the first section and the first paragraph from the word document.
	Section* sec = document->GetSections()->GetItem(0);
	Paragraph* para = sec->GetParagraphs()->GetItem(0);

	//Loop through the textranges and set CharacterFormat.Hidden property as true to hide the texts.
	for (int i = 0; i < para->GetChildObjects()->GetCount(); i++)
	{
		DocumentObject* obj = para->GetChildObjects()->GetItem(i);
		if (dynamic_cast<TextRange*>(obj) != nullptr)
		{
			TextRange* range = dynamic_cast<TextRange*>(obj);
			range->GetCharacterFormat()->SetHidden(true);
		}
	}

	//Save to file.
	document->SaveToFile(outputFile.c_str(), FileFormat::Docx2013);
	document->Close();
	delete document;
}