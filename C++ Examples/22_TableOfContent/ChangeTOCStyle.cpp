#include "pch.h"
using namespace Spire::Doc;
using namespace Spire::Common;

int main() {
	wstring input_path = DATAPATH;
	wstring inputFile = input_path + L"Template_Toc.docx";
	wstring output_path = OUTPUTPATH;
	wstring outputFile = output_path + L"ChangeTOCStyle.docx";

	//Load document from disk
	Document* doc = new Document();
	doc->LoadFromFile(inputFile.c_str());

	//Defind a Toc style
	ParagraphStyle* tocStyle = dynamic_cast<ParagraphStyle*>(Style::CreateBuiltinStyle(BuiltinStyle::Toc1, doc));
	tocStyle->GetCharacterFormat()->SetFontName(L"Aleo");
	tocStyle->GetCharacterFormat()->SetFontSize(15.0f);
	tocStyle->GetCharacterFormat()->SetTextColor(Color::GetCadetBlue());
	doc->GetStyles()->Add(tocStyle);

	//Loop through sections
	for (int i = 0; i < doc->GetSections()->GetCount(); i++)
	{
		Section* section = doc->GetSections()->GetItem(i);
		//Loop through content of section
		for (int j = 0; j < section->GetBody()->GetChildObjects()->GetCount(); j++)
		{
			DocumentObject* obj = section->GetBody()->GetChildObjects()->GetItem(j);
			//Find the structure document tag
			if (dynamic_cast<StructureDocumentTag*>(obj) != nullptr)
			{
				StructureDocumentTag* tag = dynamic_cast<StructureDocumentTag*>(obj);
				//Find the paragraph where the TOC1 locates
				for (int k = 0; k < tag->GetChildObjects()->GetCount(); k++)
				{
					DocumentObject* cObj = tag->GetChildObjects()->GetItem(k);
					if (dynamic_cast<Paragraph*>(cObj) != nullptr)
					{
						Paragraph* para = dynamic_cast<Paragraph*>(cObj);
						if (wcscmp(para->GetStyleName(), L"TOC1") == 0)
						{
							//Apply the new style for TOC1 paragraph
							para->ApplyStyle(tocStyle->GetName());
						}
					}
				}
			}
		}
	}

	//Save the Word file
	doc->SaveToFile(outputFile.c_str(), FileFormat::Docx2013);
	doc->Close();
	delete doc;
}