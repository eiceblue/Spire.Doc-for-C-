#include "pch.h"
#include <algorithm>
using namespace Spire::Doc;
using namespace Spire::Common;

class TextRangeLocation
{
private:
	TextRange* m_Text;
public:
	TextRangeLocation::TextRangeLocation(TextRange* text)
	{
		this->SetText(text);
	}

	TextRange* TextRangeLocation::GetText()
	{
		return m_Text;
	}

	void TextRangeLocation::SetText(TextRange* value)
	{
		m_Text = value;
	}

	Paragraph* TextRangeLocation::GetOwner()
	{
		return this->GetText()->GetOwnerParagraph();
	}

	int TextRangeLocation::GetIndex()
	{
		return this->GetOwner()->GetChildObjects()->IndexOf(this->GetText());
	}

	int TextRangeLocation::CompareTo(TextRangeLocation* other)
	{
		return -(this->GetIndex() - other->GetIndex());
	}
};
int main() {
	wstring input_path = DATAPATH;
	wstring inputFile = input_path + L"ReplaceWithHtml.docx";
	wstring output_path = OUTPUTPATH;
	wstring outputFile = output_path + L"ReplaceWithHtml.docx";
	wifstream input1(input_path + L"InputHtml1.txt");

	wstring HTML(istreambuf_iterator<wchar_t>(input1), {});
	//Load the document from disk.  
	Document* document = new Document();
	document->LoadFromFile(inputFile.c_str());

	//collect the objects which is used to replace text
	vector<DocumentObject*> replacement;

	//create a temporary section
	Section* tempSection = document->AddSection();

	//add a paragraph to append html
	Paragraph* par = tempSection->AddParagraph();
	par->AppendHTML(HTML.c_str());

	//get the objects in temporary section
	for (int i = 0; i < tempSection->GetBody()->GetChildObjects()->GetCount(); i++)
	{
		DocumentObject* obj = tempSection->GetBody()->GetChildObjects()->GetItem(i);
		DocumentObject* docObj = dynamic_cast<DocumentObject*>(obj);
		replacement.push_back(docObj);
	}

	//Find all text which will be replaced.
	vector<TextSelection*> selections = document->FindAllString(L"[#placeholder]", false, true);

	vector<TextRangeLocation*> locations;
	for (auto selection : selections)
	{
		/*TextRangeLocation tempVar(selection->GetAsOneRange());
		locations.push_back(&tempVar);*/
		TextRangeLocation* tempVar = new TextRangeLocation(selection->GetAsOneRange());
		locations.push_back(tempVar);
	}
	sort(locations.begin(), locations.end());

	for (auto location : locations)
	{
		//replace the text with HTML.c_str()
		ReplaceWithHTML(location, replacement);
	}

	//remove the temp section
	document->GetSections()->Remove(tempSection);

	//Save the document.
	document->SaveToFile(outputFile.c_str(), FileFormat::Docx);
	document->Close();
	delete document;
}

void ReplaceWithHTML(TextRangeLocation* location, vector<DocumentObject*>& replacement)
{
	TextRange* textRange = location->GetText();

	//textRange index
	int index = location->GetIndex();

	//get owener paragraph
	Paragraph* paragraph = location->GetOwner();

	//get owner text Body
	Body* sectionBody = paragraph->GetOwnerTextBody();

	//get the index of paragraph in section
	int paragraphIndex = sectionBody->GetChildObjects()->IndexOf(paragraph);

	int replacementIndex = -1;
	if (index == 0)
	{
		//remove the first child object
		paragraph->GetChildObjects()->RemoveAt(0);

		replacementIndex = sectionBody->GetChildObjects()->IndexOf(paragraph);
	}
	else if (index == paragraph->GetChildObjects()->GetCount() - 1)
	{
		paragraph->GetChildObjects()->RemoveAt(index);
		replacementIndex = paragraphIndex + 1;
	}
	else
	{
		//split owner paragraph
		Paragraph* paragraph1 = dynamic_cast<Paragraph*>(paragraph->Clone());
		while (paragraph->GetChildObjects()->GetCount() > index)
		{
			paragraph->GetChildObjects()->RemoveAt(index);
		}
		int i = 0;
		int count = index + 1;
		while (i < count)
		{
			paragraph1->GetChildObjects()->RemoveAt(0);
			i += 1;
		}
		sectionBody->GetChildObjects()->Insert(paragraphIndex + 1, paragraph1);

		replacementIndex = paragraphIndex + 1;
	}

	//insert replacement
	for (int i = 0; i <= replacement.size() - 1; i++)
	{
		sectionBody->GetChildObjects()->Insert(replacementIndex + i, replacement[i]->Clone());
	}
}