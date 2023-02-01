#include "pch.h"
using namespace Spire::Doc;

int main() {
	wstring input_path = DATAPATH;
	wstring inputFile = input_path + L"GetRevisions.docx";
	wstring output_path = OUTPUTPATH;
	wstring outputFile_1 = output_path + L"insertRevisions.txt";
	wstring outputFile_2 = output_path + L"deleteRevisions.txt";
	RemoveDirectoryW(outputFile_1.c_str());
	RemoveDirectoryW(outputFile_2.c_str());

	Document* document = new Document();
	document->LoadFromFile(inputFile.c_str());
	wstring* insertRevision = new wstring();
	insertRevision->append(L"Insert revisions:\n");
	int index_insertRevision = 0;
	wstring* deleteRevision = new wstring();
	deleteRevision->append(L"Delete revisions:\n");
	int index_deleteRevision = 0;
	//Traverse sections
	int sectionCount = document->GetSections()->GetCount();
	for (int i = 0; i < sectionCount; i++)
	{
		Section* sec = document->GetSections()->GetItem(i);
		//Iterate through the element under GetBody() in the section
		int secChildObjectsCount = sec->GetBody()->GetChildObjects()->GetCount();
		for (int j = 0; j < secChildObjectsCount; j++)
		{
			DocumentObject* docItem = sec->GetBody()->GetChildObjects()->GetItem(j);
			if (dynamic_cast<Paragraph*>(docItem) != nullptr)
			{
				Paragraph* para = dynamic_cast<Paragraph*>(docItem);
				//Determine if the paragraph is an insertion revision
				if (para->GetIsInsertRevision())
				{
					index_insertRevision++;
					insertRevision->append(L"Index: " + to_wstring(index_insertRevision) + L"\n");
					//Get insertion revision
					EditRevision* insRevison = para->GetInsertRevision();

					//Get insertion revision type
					EditRevisionType insType = insRevison->GetType();
					insertRevision->append(L"Type: " + getRevisionType(insType) + L"\n");
					//Get insertion revision author
					wstring insAuthor = insRevison->GetAuthor();
					insertRevision->append(L"Author: " + insAuthor + L"\n");
				}
				//Determine if the paragraph is a delete revision
				else if (para->GetIsDeleteRevision())
				{
					index_deleteRevision++;
					deleteRevision->append(L"Index: " + to_wstring(index_deleteRevision) + L"\n");
					EditRevision* delRevison = para->GetDeleteRevision();
					EditRevisionType delType = delRevison->GetType();
					deleteRevision->append(L"Type: " + getRevisionType(delType) + L"\n");
					wstring delAuthor = delRevison->GetAuthor();
					deleteRevision->append(L"Author: " + delAuthor + L"\n");
				}
				//Iterate through the element in the paragraph
				int paraChildObjectsCount = para->GetChildObjects()->GetCount();
				for (int k = 0; k < paraChildObjectsCount; k++)
				{
					DocumentObject* obj = para->GetChildObjects()->GetItem(k);
					if (dynamic_cast<TextRange*>(obj) != nullptr)
					{
						TextRange* textRange = dynamic_cast<TextRange*>(obj);
						//Determine if the textrange is an insertion revision
						if (textRange->GetIsInsertRevision())
						{
							index_insertRevision++;
							insertRevision->append(L"Index: " + to_wstring(index_insertRevision) + L"\n");
							EditRevision* insRevison = textRange->GetInsertRevision();
							EditRevisionType insType = insRevison->GetType();
							insertRevision->append(L"Type: " + getRevisionType(insType) + L"\n");
							wstring insAuthor = insRevison->GetAuthor();
							insertRevision->append(L"Author: " + insAuthor + L"\n");
						}
						else if (textRange->GetIsDeleteRevision())
						{
							index_deleteRevision++;
							deleteRevision->append(L"Index: " + to_wstring(index_deleteRevision) + L"\n");
							//Determine if the textrange is a delete revision
							EditRevision* delRevison = textRange->GetDeleteRevision();
							EditRevisionType delType = delRevison->GetType();
							deleteRevision->append(L"Type: " + getRevisionType(delType) + L"\n");
							wstring delAuthor = delRevison->GetAuthor();
							deleteRevision->append(L"Author: " + delAuthor + L"\n");
						}
					}
				}
			}
		}
	}
	wofstream out1;
	out1.open(outputFile_1.c_str());
	out1.flush();
	out1 << insertRevision->c_str();
	out1.close();

	wofstream out2;
	out2.open(outputFile_2.c_str());
	out2.flush();
	out2 << deleteRevision->c_str();
	out2.close();

	delete deleteRevision;
	delete insertRevision;
}

wstring getRevisionType(EditRevisionType type)
{
	switch (type)
	{
	case EditRevisionType::Deletion:
		return L"Deletion";
		break;
	case EditRevisionType::Insertion:
		return L"Insertion";
		break;
	}
}