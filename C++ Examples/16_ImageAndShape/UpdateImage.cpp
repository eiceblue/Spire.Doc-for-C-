#include "pch.h"
using namespace Spire::Doc;

int main() {
	wstring input_path = DATAPATH;
	wstring inputFile = input_path  + L"ImageTemplate.docx";
	wstring imagePath = input_path + L"E-iceblue.png";
	wstring output_path = OUTPUTPATH;
	wstring outputFile = output_path + L"UpdateImage.docx";

	//Load Document
	Document* doc = new Document();
	doc->LoadFromFile(inputFile.c_str());

	//Get all pictures in the Word document
	vector<DocumentObject*> pictures;
	for (int i = 0; i < doc->GetSections()->GetCount(); i++)
	{
		Section* sec = doc->GetSections()->GetItem(i);
		for (int j = 0; j < sec->GetParagraphs()->GetCount(); j++)
		{
			Paragraph* para = sec->GetParagraphs()->GetItem(j);
			for (int k = 0; k < para->GetChildObjects()->GetCount(); k++)
			{
				DocumentObject* docObj = para->GetChildObjects()->GetItem(k);
				if (docObj->GetDocumentObjectType() == DocumentObjectType::Picture)
				{
					pictures.push_back(docObj);
				}
			}
		}
	}

	//Replace the first picture with a new image file
	DocPicture* picture = dynamic_cast<DocPicture*>(pictures[0]);
	picture->LoadImageSpire(imagePath.c_str());

	//Save and launch document
	doc->SaveToFile(outputFile.c_str(), FileFormat::Docx);
	doc->Close();
	delete doc;
}
