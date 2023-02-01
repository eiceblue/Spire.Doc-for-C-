#include "pch.h"
using namespace Spire::Doc;

int main() {
	wstring input_path = DATAPATH;
	wstring inputFile = input_path  + L"ShapeWithAlternativeText.docx";
	wstring output_path = OUTPUTPATH;
	wstring outputFile = output_path + L"GetAlternativeText.txt";

	//Create a document
	Document* document = new Document();
	//Create string builder
	wstring* builder = new wstring();
	document->LoadFromFile(inputFile.c_str());

	//Loop through shapes and get the AlternativeText
	for (int i = 0; i < document->GetSections()->GetCount(); i++)
	{
		Section* section = document->GetSections()->GetItem(i);
		for (int j = 0; j < section->GetParagraphs()->GetCount(); j++)
		{
			Paragraph* para = section->GetParagraphs()->GetItem(j);
			for (int k = 0; k < para->GetChildObjects()->GetCount(); k++)
			{
				DocumentObject* obj = para->GetChildObjects()->GetItem(k);
				if (dynamic_cast<ShapeObject*>(obj) != nullptr)
				{
					wstring text = (dynamic_cast<ShapeObject*>(obj))->GetAlternativeText();
					//Append the alternative text in builder
					builder->append(text);
					builder->append(L"\n");
				}
			}
		}
	}

	//Save doc file.
	wofstream write(outputFile);
	write << builder->c_str();
	write.close();
	document->Close();
	delete document;
	delete builder;
}
