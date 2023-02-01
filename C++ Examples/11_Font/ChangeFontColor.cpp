#include "pch.h"
using namespace Spire::Doc;
using namespace Spire::Common;

int main() {
	wstring input_path = DATAPATH;
	wstring inputFile = input_path  + L"Sample.docx";
	wstring output_path = OUTPUTPATH;
	wstring outputFile = output_path + L"ChangeFontColor.docx";

	//Load the document
	Document* doc = new Document();
	doc->LoadFromFile(inputFile.c_str());

	//Get the first section and first paragraph
	Section* section = doc->GetSections()->GetItem(0);

	Paragraph* p1 = section->GetParagraphs()->GetItem(0);
	//Iterate through the childObjects of the paragraph 1 
	int p1ChildObjectsCount = p1->GetChildObjects()->GetCount();
	for (int i = 0; i < p1ChildObjectsCount; i++)
	{
		DocumentObject* childObj = p1->GetChildObjects()->GetItem(i);
		if (dynamic_cast<TextRange*>(childObj) != nullptr)
		{
			//Change text color
			TextRange* tr = dynamic_cast<TextRange*>(childObj);
			tr->GetCharacterFormat()->SetTextColor(Color::GetRosyBrown());
		}
	}

	//Get the second paragraph
	Paragraph* p2 = section->GetParagraphs()->GetItem(1);

	//Iterate through the childObjects of the paragraph 2
	int p2ChildObjectsCount = p2->GetChildObjects()->GetCount();
	for (int i = 0; i < p2ChildObjectsCount; i++)
	{
		DocumentObject* childObj = p2->GetChildObjects()->GetItem(i);
		if (dynamic_cast<TextRange*>(childObj) != nullptr)
		{
			//Change text color
			TextRange* tr = dynamic_cast<TextRange*>(childObj);
			tr->GetCharacterFormat()->SetTextColor(Color::GetDarkGreen());
		}
	}

	//Save and launch document
	doc->SaveToFile(outputFile.c_str(), FileFormat::Docx);
	doc->Close();
	delete doc;
}
