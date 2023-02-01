#include "pch.h"
using namespace Spire::Doc;

int main() {
	wstring input_path = DATAPATH;
	wstring inputFile = input_path + L"BetweenParagraphStyle.docx";
	wstring output_path = OUTPUTPATH;
	wstring outputFile = output_path + L"BetweenParagraphStyles.docx";

	//Create the first document
	Document* sourceDocument = new Document();

	//Load the source document from disk.
	sourceDocument->LoadFromFile(inputFile.c_str());

	//Create a destination document
	Document* destinationDoc = new Document();

	//Add a section
	Section* section = destinationDoc->AddSection();

	//Extract content between the first paragraph to the third paragraph
	ExtractBetweenParagraphStyles(sourceDocument, destinationDoc, L"1", L"2");

	//Save the document.
	destinationDoc->SaveToFile(outputFile.c_str(), FileFormat::Docx);
	sourceDocument->Close();
	destinationDoc->Close();
	delete sourceDocument;
	delete destinationDoc;
}

void ExtractBetweenParagraphStyles(Document* sourceDocument, Document* destinationDocument, const wstring& stylename1, const wstring& stylename2)
{
	int startindex = 0;
	int endindex = 0;
	//travel the sections of source document

	for (int i = 0; i < sourceDocument->GetSections()->GetCount(); i++)
	{
		Section* section = sourceDocument->GetSections()->GetItem(i);
		//travel the paragraphs
		for (int j = 0; j < section->GetParagraphs()->GetCount(); j++)
		{
			Paragraph* paragraph = section->GetParagraphs()->GetItem(j);
			//Judge paragraph style1
			if (paragraph->GetStyleName() == stylename1)
			{
				//Get the paragraph index
				startindex = section->GetBody()->GetParagraphs()->IndexOf(paragraph);
			}
			//Judge paragraph style2
			if (paragraph->GetStyleName() == stylename2)
			{
				//Get the paragraph index
				endindex = section->GetBody()->GetParagraphs()->IndexOf(paragraph);
			}
		}
		//Extract the content
		for (int i = startindex + 1; i < endindex; i++)
		{
			//Clone the ChildObjects of source document
			DocumentObject* doobj = sourceDocument->GetSections()->GetItem(0)->GetBody()->GetChildObjects()->GetItem(i)->Clone();

			//Add to destination document 
			destinationDocument->GetSections()->GetItem(0)->GetBody()->GetChildObjects()->Add(doobj);
		}
	}
}