#include "pch.h"
using namespace Spire::Doc;

int main() {
	wstring input_path = DATAPATH;
	wstring inputFile_1 = input_path + L"ReplaceContentWithDoc.docx";
	wstring inputFile_2 = input_path + L"Insert.docx";
	wstring output_path = OUTPUTPATH;
	wstring outputFile = output_path + L"ReplaceContentWithDoc.docx";

	//Create the first document
	Document* document1 = new Document();

	//Load the first document from disk.
	document1->LoadFromFile(inputFile_1.c_str());

	//Create the second document
	Document* document2 = new Document();

	//Load the second document from disk.
	document2->LoadFromFile(inputFile_2.c_str());

	//Get the first section of the first document 
	Section* section1 = document1->GetSections()->GetItem(0);

	//Create a regex
	Regex* regex = new Regex(L"\\[MY_DOCUMENT\]", RegexOptions::None);

	//Find the text by regex
	vector<TextSelection*> textSections = document1->FindAllPattern(regex);

	//Travel the found strings
	for (auto seletion : textSections)
	{

		//Get the para
		Paragraph* para = seletion->GetAsOneRange()->GetOwnerParagraph();

		//Get textRange
		TextRange* textRange = seletion->GetAsOneRange();

		//Get the para index
		int index = section1->GetBody()->GetChildObjects()->IndexOf(para);

		//Insert the paragraphs of document2
		for (int i = 0; i < document2->GetSections()->GetCount(); i++)
		{
			Section* section2 = document2->GetSections()->GetItem(i);
			for (int j = 0; j < section2->GetParagraphs()->GetCount(); j++)
			{
				Paragraph* paragraph = section2->GetParagraphs()->GetItem(j);
				section1->GetBody()->GetChildObjects()->Insert(index, dynamic_cast<Paragraph*>(paragraph->Clone()));
			}
		}
		//Remove the found textRange
		para->GetChildObjects()->Remove(textRange);
	}

	//Save the document.
	document1->SaveToFile(outputFile.c_str(), FileFormat::Docx);
	document1->Dispose();
	document2->Dispose();
	delete document1;
	delete document2;
}