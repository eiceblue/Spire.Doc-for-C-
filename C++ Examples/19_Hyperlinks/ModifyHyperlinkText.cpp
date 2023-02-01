#include "pch.h"
using namespace Spire::Doc;

int main() {
	wstring input_path = DATAPATH;
	wstring inputFile = input_path + L"Hyperlinks.docx";
	wstring output_path = OUTPUTPATH;
	wstring outputFile = output_path + L"ModifyHyperlinkText.docx";

	//Load Document
	Document* doc = new Document();
	doc->LoadFromFile(inputFile.c_str());

	vector<Field*> hyperlinks;
	for (int i = 0; i < doc->GetSections()->GetCount(); i++)
	{
		Section* section = doc->GetSections()->GetItem(i);
		for (int j = 0; j < section->GetBody()->GetChildObjects()->GetCount(); j++)
		{
			DocumentObject* docObj = section->GetBody()->GetChildObjects()->GetItem(j);
			if (docObj->GetDocumentObjectType() == DocumentObjectType::Paragraph)
			{
				Paragraph* para = dynamic_cast<Paragraph*>(docObj);
				for (int k = 0; k < para->GetChildObjects()->GetCount(); k++)
				{
					DocumentObject* obj = para->GetChildObjects()->GetItem(k);
					if (obj->GetDocumentObjectType() == DocumentObjectType::Field)
					{
						Field* field = dynamic_cast<Field*>(obj);
						if (field->GetType() == FieldType::FieldHyperlink)
						{
							hyperlinks.push_back(field);
						}
					}
				}
			}
		}
	}

	//Reset the property of hyperlinks[0].FieldText by using the index of the hyperlink
	hyperlinks[0]->SetFieldText(L"Spire.Doc component");

	//Save and launch document
	doc->SaveToFile(outputFile.c_str(), FileFormat::Docx);
	doc->Close();
	delete doc;
}
