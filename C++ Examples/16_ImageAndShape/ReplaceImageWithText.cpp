#include "pch.h"
using namespace Spire::Doc;

int main() {
	wstring input_path = DATAPATH;
	wstring inputFile = input_path  + L"ImageTemplate.docx";
	wstring output_path = OUTPUTPATH;
	wstring outputFile = output_path + L"ReplaceImageWithText.docx";

	//Load Document
	Document* doc = new Document();
	doc->LoadFromFile(inputFile.c_str());

	//Replace all pictures with texts
	int j = 1;
	for (int i = 0; i < doc->GetSections()->GetCount(); i++)
	{
		Section* sec = doc->GetSections()->GetItem(i);
		for (int k = 0; k < sec->GetParagraphs()->GetCount(); k++)
		{
			Paragraph* para = sec->GetParagraphs()->GetItem(k);
			vector<DocumentObject*> pictures;
			//Get all pictures in the Word document
			for (int m = 0; m < para->GetChildObjects()->GetCount(); m++)
			{
				DocumentObject* docObj = para->GetChildObjects()->GetItem(m);
				if (docObj->GetDocumentObjectType() == DocumentObjectType::Picture)
				{
					pictures.push_back(docObj);
				}
			}

			//Replace pitures with the text "Here was image {image index}"
			for (auto pic : pictures)
			{
				int index = para->GetChildObjects()->IndexOf(pic);
				TextRange* range = new TextRange(doc);
				wstring temp= L"Here was image " + to_wstring(j) + L"";
				range->SetText(temp.c_str());
				para->GetChildObjects()->Insert(index, range);
				para->GetChildObjects()->Remove(pic);
				j++;
			}
		}
	}

	//Save and launch document
	doc->SaveToFile(outputFile.c_str(), FileFormat::Docx);
	doc->Close();
	delete doc;
}
