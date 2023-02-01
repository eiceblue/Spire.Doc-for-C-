#include "pch.h"
using namespace Spire::Doc;

int main() {
	wstring input_path = DATAPATH;
	wstring inputFile1 = input_path + L"ResetPageNumber1.docx";
	wstring inputFile2 = input_path + L"ResetPageNumber2.docx";
	wstring inputFile3 = input_path + L"ResetPageNumber3.docx";
	wstring output_path = OUTPUTPATH;
	wstring outputFile = output_path + L"ResetPageNumber.docx";

	//Create three Word documents and load three different Word documents from disk.
	Document* document1 = new Document();
	document1->LoadFromFile(inputFile1.c_str());

	Document* document2 = new Document();
	document2->LoadFromFile(inputFile2.c_str());

	Document* document3 = new Document();
	document3->LoadFromFile(inputFile3.c_str());

	//Use section method to combine all documents into one word document.
	for (int i = 0; i < document2->GetSections()->GetCount(); i++)
	{
		Section* sec = document2->GetSections()->GetItem(i);
		document1->GetSections()->Add(sec->Clone());
	}
	for (int j = 0; j < document3->GetSections()->GetCount(); j++)
	{
		Section* sec = document3->GetSections()->GetItem(j);
		document1->GetSections()->Add(sec->Clone());
	}

	//Traverse every section of document1.
	for (int k = 0; k < document1->GetSections()->GetCount(); k++)
	{
		Section* sec = document1->GetSections()->GetItem(k);
		//Traverse every object of the footer.
		for (int m = 0; m < sec->GetHeadersFooters()->GetFooter()->GetChildObjects()->GetCount(); m++)
		{
			DocumentObject* obj = sec->GetHeadersFooters()->GetFooter()->GetChildObjects()->GetItem(m);
			if (obj->GetDocumentObjectType() == DocumentObjectType::StructureDocumentTag)
			{
				DocumentObject* para = obj->GetChildObjects()->GetItem(m);
				for (int n = 0; n < para->GetChildObjects()->GetCount(); n++)
				{
					DocumentObject* item = para->GetChildObjects()->GetItem(n);
					if (item->GetDocumentObjectType() == DocumentObjectType::Field)
					{
						//Find the item and its field type is FieldNumPages.
						if ((dynamic_cast<Field*>(item))->GetType() == FieldType::FieldNumPages)
						{
							//Change field type to FieldSectionPages.
							(dynamic_cast<Field*>(item))->SetType(FieldType::FieldSectionPages);
						}
					}
				}
			}
		}
	}

	//Restart page number of section and set the starting page number to 1.
	document1->GetSections()->GetItem(1)->GetPageSetup()->SetRestartPageNumbering(true);
	document1->GetSections()->GetItem(1)->GetPageSetup()->SetPageStartingNumber(1);

	document1->GetSections()->GetItem(2)->GetPageSetup()->SetRestartPageNumbering(true);
	document1->GetSections()->GetItem(2)->GetPageSetup()->SetPageStartingNumber(1);

	//Save to file.
	document1->SaveToFile(outputFile.c_str(), FileFormat::Docx2013);
	document1->Close();
	delete document1;
	delete document2;
	delete document3;
}
