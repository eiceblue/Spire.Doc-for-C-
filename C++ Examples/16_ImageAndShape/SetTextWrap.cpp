#include "pch.h"
using namespace Spire::Doc;

int main() {
	wstring input_path = DATAPATH;
	wstring inputFile = input_path  + L"ImageTemplate.docx";
	wstring output_path = OUTPUTPATH;
	wstring outputFile = output_path + L"SetTextWrap.docx";

	//Load Document
	Document* doc = new Document();
	doc->LoadFromFile(inputFile.c_str());

	for (int i = 0; i < doc->GetSections()->GetCount(); i++)
	{
		Section* sec = doc->GetSections()->GetItem(i);
		for (int j = 0; j < sec->GetParagraphs()->GetCount(); j++)
		{
			Paragraph* para = sec->GetParagraphs()->GetItem(j);
			vector<DocumentObject*> pictures;
			//Get all pictures in the Word document
			for (int k = 0; k < para->GetChildObjects()->GetCount(); k++)
			{
				DocumentObject* docObj = para->GetChildObjects()->GetItem(k);
				if (docObj->GetDocumentObjectType() == DocumentObjectType::Picture)
				{
					pictures.push_back(docObj);
				}
			}

			//Set text wrap styles for each piture
			for (auto pic : pictures)
			{
				DocPicture* picture = dynamic_cast<DocPicture*>(pic);
				picture->SetTextWrappingStyle(TextWrappingStyle::Through);
				picture->SetTextWrappingType(TextWrappingType::Both);
			}
		}
	}

	//Save and launch document
	doc->SaveToFile(outputFile.c_str(), FileFormat::Docx);
	doc->Close();
	delete doc;
}
