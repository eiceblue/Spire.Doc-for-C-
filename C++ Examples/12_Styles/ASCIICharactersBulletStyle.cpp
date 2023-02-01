#include "pch.h"
using namespace Spire::Doc;

int main() {
	wstring output_path = OUTPUTPATH;
	wstring outputFile = output_path + L"ASCIICharactersBulletStyle.docx";

	//Create a new document
	Document* document = new Document();
	Section* section = document->AddSection();

	//Create four list styles based on different ASCII characters
	ListStyle* listStyle1 = new ListStyle(document, ListType::Bulleted);
	listStyle1->SetName(L"liststyle");
	listStyle1->GetLevels()->GetItem(0)->SetBulletCharacter(L"\x006e");
	listStyle1->GetLevels()->GetItem(0)->GetCharacterFormat()->SetFontName(L"Wingdings");
	document->GetListStyles()->Add(listStyle1);

	ListStyle* listStyle2 = new ListStyle(document, ListType::Bulleted);
	listStyle2->SetName(L"liststyle2");
	listStyle2->GetLevels()->GetItem(0)->SetBulletCharacter(L"\x0075");
	listStyle2->GetLevels()->GetItem(0)->GetCharacterFormat()->SetFontName(L"Wingdings");
	document->GetListStyles()->Add(listStyle2);

	ListStyle* listStyle3 = new ListStyle(document, ListType::Bulleted);
	listStyle3->SetName(L"liststyle3");
	listStyle3->GetLevels()->GetItem(0)->SetBulletCharacter(L"\x00b2");
	listStyle3->GetLevels()->GetItem(0)->GetCharacterFormat()->SetFontName(L"Wingdings");
	document->GetListStyles()->Add(listStyle3);

	ListStyle* listStyle4 = new ListStyle(document, ListType::Bulleted);
	listStyle4->SetName(L"liststyle4");
	listStyle4->GetLevels()->GetItem(0)->SetBulletCharacter(L"\x00d8");
	listStyle4->GetLevels()->GetItem(0)->GetCharacterFormat()->SetFontName(L"Wingdings");
	document->GetListStyles()->Add(listStyle4);

	//Add four paragraphs and apply list style separately
	Paragraph* p1 = section->GetBody()->AddParagraph();
	p1->AppendText(L"Spire.Doc for C++");
	p1->GetListFormat()->ApplyStyle(listStyle1->GetName());
	p1->GetListFormat()->ApplyStyle(listStyle1->GetName());

	Paragraph* p2 = section->GetBody()->AddParagraph();
	p2->AppendText(L"Spire.Doc for C++");
	p2->GetListFormat()->ApplyStyle(listStyle2->GetName());

	Paragraph* p3 = section->GetBody()->AddParagraph();
	p3->AppendText(L"Spire.Doc for C++");
	p3->GetListFormat()->ApplyStyle(listStyle3->GetName());

	Paragraph* p4 = section->GetBody()->AddParagraph();
	p4->AppendText(L"Spire.Doc for C++");
	p4->GetListFormat()->ApplyStyle(listStyle4->GetName());

	//Save to docx file.
	document->SaveToFile(outputFile.c_str(), FileFormat::Docx);
	document->Close();
	delete document;
}
