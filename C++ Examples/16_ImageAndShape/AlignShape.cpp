#include "pch.h"
using namespace Spire::Doc;

int main() {
	wstring input_path = DATAPATH;
	wstring inputFile = input_path  + L"Shapes.docx";
	wstring output_path = OUTPUTPATH;
	wstring outputFile = output_path + L"AlignShape.docx";

	//Load Document
	Document* doc = new Document();
	doc->LoadFromFile(inputFile.c_str());

	Section* section = doc->GetSections()->GetItem(0);

	for (int i = 0; i < section->GetParagraphs()->GetCount(); i++)
	{
		Paragraph* para = section->GetParagraphs()->GetItem(i);
		for (int j = 0; j < para->GetChildObjects()->GetCount(); j++)
		{
			DocumentObject* obj = para->GetChildObjects()->GetItem(j);
			if (dynamic_cast<ShapeObject*>(obj) != nullptr)
			{
				//Set the horizontal alignment as center
				(dynamic_cast<ShapeObject*>(obj))->SetHorizontalAlignment(ShapeHorizontalAlignment::Center);
			}
		}
	}

	//Save and launch document
	doc->SaveToFile(outputFile.c_str(), FileFormat::Docx);
	doc->Close();
	delete doc;
}
