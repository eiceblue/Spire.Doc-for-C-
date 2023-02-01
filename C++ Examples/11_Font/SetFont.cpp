#include "pch.h"
using namespace Spire::Doc;

int main() {
	wstring input_path = DATAPATH;
	wstring inputFile = input_path  + L"Sample.docx";
	wstring output_path = OUTPUTPATH;
	wstring outputFile = output_path + L"SetFont.docx";
	
	//Load the document
	Document* doc = new Document();
	doc->LoadFromFile(inputFile.c_str());

	//Get the first section 
	Section* s = doc->GetSections()->GetItem(0);

	//Get the second paragraph
	Paragraph* p = s->GetParagraphs()->GetItem(1);

	//Create a characterFormat object
	CharacterFormat* format = new CharacterFormat(doc);
	//Set font
	format->SetFontName(L"Arial");
	format->SetFontSize(16);

	//Loop through the childObjects of paragraph 
	int pChildObjectsCount = p->GetChildObjects()->GetCount();
	for (int i = 0; i < pChildObjectsCount; i++)
	{
		DocumentObject* childObj = p->GetChildObjects()->GetItem(i);
		if (dynamic_cast<TextRange*>(childObj) != nullptr)
		{
			//Apply character format
			TextRange* tr = dynamic_cast<TextRange*>(childObj);
			tr->ApplyCharacterFormat(format);
		}
	}

	//Save and launch document
	doc->SaveToFile(outputFile.c_str(), FileFormat::Docx);
	doc->Close();
	delete doc;
}	