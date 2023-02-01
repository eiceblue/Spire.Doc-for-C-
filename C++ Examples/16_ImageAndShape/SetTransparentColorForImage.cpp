#include "pch.h"
using namespace Spire::Doc;
using namespace Spire::Common;

int main() {
	wstring input_path = DATAPATH;
	wstring inputFile = input_path  + L"ImageTemplate.docx";
	wstring output_path = OUTPUTPATH;
	wstring outputFile = output_path + L"SetTransparentColorForImage.docx";

	//Load Document
	Document* doc = new Document();
	doc->LoadFromFile(inputFile.c_str());

	//Get the first paragraph in the first section
	Paragraph* paragraph = doc->GetSections()->GetItem(0)->GetParagraphs()->GetItem(0);

	//Set the blue color of the image(s) in the paragraph to transperant
	for (int i = 0; i < paragraph->GetChildObjects()->GetCount(); i++)
	{
		DocumentObject* obj = paragraph->GetChildObjects()->GetItem(i);
		if (dynamic_cast<DocPicture*>(obj) != nullptr)
		{
			DocPicture* picture = dynamic_cast<DocPicture*>(obj);
			picture->SetTransparentColor(Color::GetBlue());
		}
	}

	//Save and launch document
	doc->SaveToFile(outputFile.c_str(), FileFormat::Docx);
	doc->Close();
	delete doc;
}
