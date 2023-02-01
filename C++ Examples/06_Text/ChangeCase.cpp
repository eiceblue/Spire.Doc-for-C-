#include "pch.h"
using namespace Spire::Doc;

int main() {
	wstring input_path = DATAPATH;
	wstring inputFile = input_path + L"Text1.docx";
	wstring output_path = OUTPUTPATH;
	wstring outputFile = output_path + L"ChangeCase.docx";

	// Create a new document and load from file;
	Document* doc = new Document();
	doc->LoadFromFile(inputFile.c_str());
	TextRange* textRange;
	//Get the first paragraph and set its CharacterFormat to AllCaps
	Paragraph* para1 = doc->GetSections()->GetItem(0)->GetParagraphs()->GetItem(1);

	for (int i = 0; i < para1->GetChildObjects()->GetCount(); i++)
	{
		DocumentObject* obj = para1->GetChildObjects()->GetItem(i);
		if (dynamic_cast<TextRange*>(obj) != nullptr)
		{
			textRange = dynamic_cast<TextRange*>(obj);
			textRange->GetCharacterFormat()->SetAllCaps(true);
		}
	}

	//Get the third paragraph and set its CharacterFormat to IsSmallCaps
	Paragraph* para2 = doc->GetSections()->GetItem(0)->GetParagraphs()->GetItem(3);
	for (int i = 0; i < para2->GetChildObjects()->GetCount(); i++)
	{
		DocumentObject* obj = para2->GetChildObjects()->GetItem(i);
		if (dynamic_cast<TextRange*>(obj) != nullptr)
		{
			textRange = dynamic_cast<TextRange*>(obj);
			textRange->GetCharacterFormat()->SetIsSmallCaps(true);
		}
	}

	//Save and launch the document
	doc->SaveToFile(outputFile.c_str(), FileFormat::Docx2013);
	doc->Close();
	delete doc;
}