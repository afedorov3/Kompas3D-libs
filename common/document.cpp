#include "stdafx.h"
#include "document.h"

BOOL Document::IsDocumentValid( ksAPI7::IKompasDocument3DPtr & doc3d )
{
	ksAPI7::IFeature7Ptr feature(doc3d->TopPart);
	if (feature == NULL)
		return FALSE;

	return feature->Valid;
}

BOOL Document::RebuildDocument( ksAPI7::IKompasDocument3DPtr & doc3d )
{
	return doc3d->RebuildDocument();
}

BOOL Document::RebuildModelEx( ksAPI7::IKompasDocument3DPtr & doc3d, BOOL redraw )
{
	// document 3D API7 to API5
	IDocument3DPtr doc3d5( IUnknownPtr(ksTransferInterface( doc3d, ksAPI3DCom, 0/*any document*/ ), false/*don't AddRef*/) );
	if (doc3d5 == NULL)
		return FALSE;

	IPartPtr part( doc3d5->GetPart( pTop_Part ), false/*AddRef*/ );
	if (part  == NULL) 
		return FALSE;

	return part->RebuildModelEx(redraw);
}

LONG Document::EmbodimentCount( ksAPI7::IKompasDocument3DPtr & doc3d )
{
	ksAPI7::IEmbodimentsManagerPtr em(doc3d);
	if (em == NULL)
		return -1;

	return em->GetEmbodimentCount();
}

ksAPI7::IEmbodimentPtr Document::EmbodimentPtrGet( ksAPI7::IKompasDocument3DPtr & doc3d, LONG index )
{
	ksAPI7::IEmbodimentsManagerPtr em(doc3d);
	if (em == NULL)
		return NULL;

	if (index >= 0) {
		if (index < em->EmbodimentCount)
			return em->GetEmbodiment(index);
		else
			return NULL;
	} else
		return em->CurrentEmbodiment;

	return NULL;
}

BOOL Document::EmbodimentGet( ksAPI7::IKompasDocument3DPtr & doc3d, LONG & index )
{
	ksAPI7::IEmbodimentsManagerPtr em(doc3d);
	if (em == NULL)
		return FALSE;

	index = em->GetCurrentEmbodimentIndex();
	return TRUE;
}

BOOL Document::EmbodimentSet( ksAPI7::IKompasDocument3DPtr & doc3d, LONG index )
{
	if (index < 0)
		return FALSE;

	ksAPI7::IEmbodimentsManagerPtr em(doc3d);
	if (em == NULL)
		return FALSE;

	if ((index < 0) || (index >= em->GetEmbodimentCount()))
		return FALSE;
	em->SetCurrentEmbodiment(index);
	if (em->GetCurrentEmbodimentIndex() != index)
		return FALSE;

	return TRUE;
}

BOOL Document::EmbodimentMarkingGet( ksAPI7::IKompasDocument3DPtr & doc3d, LONG index, LONG type, CString * name, BOOL append )
{
	ksAPI7::IEmbodimentPtr e = EmbodimentPtrGet(doc3d, index);
	if (e == NULL)
		return FALSE;

	BSTR emname = e->GetMarking((ksVariantMarkingTypeEnum)type, FALSE);
	if (emname == NULL)
		return FALSE;

	if (append)
		name->Append(emname);
	else
		name->SetString(emname);

	return TRUE;
}

BOOL Document::EmbodimentNameEnum( ksAPI7::IKompasDocument3DPtr & doc3d, EMBODIMENTNAMES & names )
{
	ksAPI7::IEmbodimentsManagerPtr em(doc3d);
	if (em == NULL)
		return FALSE;

	LONG count = em->EmbodimentCount;
	names.clear();
	for(LONG i = 0; i < count; i++) {
		ksAPI7::IEmbodimentPtr e = em->Embodiment[i];
		if (e == NULL)
			return FALSE;

		EMBODIMENTNAME name;
		names.push_back(name);
		EMBODIMENTNAME *pname = &names.back();
		pname->name = (LPCWSTR)e->Name;
		pname->number = (LPWSTR)e->GetMarking(ksVMFullMarking, FALSE);
	}

	return TRUE;
}

BOOL Document::ModelEmbodimentNameGet( ksAPI7::IKompasDocument3DPtr & doc3d, LONG embodiment, LPWSTR & modelname )
{
	ksAPI7::IEmbodimentPtr e = EmbodimentPtrGet(doc3d, embodiment);
	if (e == NULL)
		return FALSE;

	modelname = e->Name;

	return TRUE;
}

BOOL Document::ModelEmbodimentNameSet( ksAPI7::IKompasDocument3DPtr & doc3d, LONG embodiment, LPCWSTR modelname )
{
	ksAPI7::IEmbodimentPtr e = EmbodimentPtrGet(doc3d, embodiment);
	if (e == NULL)
		return FALSE;

	if (wcscmp(e->Name, modelname) == 0)	 // Update() will fail when setting the same name
		return TRUE;

	e->Name = modelname;

	return e->Update();
}

BOOL Document::ModelEmbodimentNameSetFmt( ksAPI7::IKompasDocument3DPtr & doc3d, LONG embodiment, LPCWSTR fmt, ... )
{
	va_list args;
	CString str;

	va_start(args, fmt);
	str.FormatV(fmt, args);
	va_end(args);

	return ModelEmbodimentNameSet(doc3d, embodiment, (LPCWSTR)str);
}


BOOL Document::ModelNameGet( ksAPI7::IKompasDocument3DPtr & doc3d, LPWSTR & modelname )
{
	return ModelEmbodimentNameGet(doc3d, EMBODIMENT_CURRENT, modelname);
}

BOOL Document::ModelNameSet( ksAPI7::IKompasDocument3DPtr & doc3d, LPCWSTR modelname )
{
	return ModelEmbodimentNameSet(doc3d, EMBODIMENT_CURRENT, modelname);
}

BOOL Document::ModelNameSetFmt( ksAPI7::IKompasDocument3DPtr & doc3d, LPCWSTR fmt, ... )
{
	va_list args;
	CString str;

	va_start(args, fmt);
	str.FormatV(fmt, args);
	va_end(args);

	return ModelEmbodimentNameSet(doc3d, EMBODIMENT_CURRENT, (LPCWSTR)str);
}

BOOL Document::IsVariableNameCharValid( WCHAR ch, BOOL first )
{
	if (first)
		return (ch == '_' || (ch >= 'A' && ch <= 'Z') || (ch >= 'a' && ch <= 'z'));
	return (ch == '_' || (ch >= 'A' && ch <= 'Z') || (ch >= 'a' && ch <= 'z') || iswdigit(ch));
}

BOOL Document::IsVariableNameValid(LPCWSTR name, LONG len)
{
	WCHAR ch = *name;
	const LONG maxlen = 512;

	if (len == 0 || len > maxlen || (ch != '_' && !(ch >= 'A' && ch <= 'Z') && !(ch >= 'a' && ch <= 'z')))
		return FALSE;
	if (len < 0) len = -1; // length is undefined, count it
	while (--len != 0 && (ch = *++name))
		if (len < -maxlen || (ch != '_' && !(ch >= 'A' && ch <= 'Z') && !(ch >= 'a' && ch <= 'z') && !(ch >= '0' && ch <= '9')))
			return FALSE;
	return TRUE;
}

ksAPI7::IVariable7Ptr Document::EmbodimentVariable( ksAPI7::IKompasDocument3DPtr & doc3d, LPCWSTR varname, LONG embodiment, BOOL create )
{
	ksAPI7::IEmbodimentPtr e = EmbodimentPtrGet(doc3d, embodiment);
	if (e == NULL)
		return NULL;
	ksAPI7::IPart7Ptr epart = e->GetPart();
	if (epart == NULL)
		return NULL;
	ksAPI7::IFeature7Ptr feature(epart);
	if (feature == NULL)
		return NULL;
	ksAPI7::IVariable7Ptr var = feature->GetVariable(FALSE, FALSE, varname);
	if (var == NULL) {
		if (!create)
			return NULL;
		var = epart->AddVariable(varname, 1.0, L"");
	}

	return var;
}

BOOL Document::VariableValueGet( ksAPI7::IKompasDocument3DPtr & doc3d, LPCWSTR varname, DOUBLE & value, LONG embodiment )
{
	ksAPI7::IVariable7Ptr var = NULL;
	if ((var = EmbodimentVariable( doc3d, varname, embodiment)) == NULL)
		return FALSE;

	value = var->Value;

	return TRUE;
}

BOOL Document::VariableValueSet( ksAPI7::IKompasDocument3DPtr & doc3d, LPCWSTR varname, DOUBLE value, LONG embodiment )
{
	ksAPI7::IVariable7Ptr var = NULL;
	if ((var = EmbodimentVariable( doc3d, varname, embodiment)) == NULL)
		return FALSE;

	var->Value = value;

	return TRUE;
}

BOOL Document::VariableExprGet( ksAPI7::IKompasDocument3DPtr & doc3d, LPCWSTR varname, LPCWSTR & expr, LONG embodiment )
{
	ksAPI7::IVariable7Ptr var = NULL;
	if ((var = EmbodimentVariable( doc3d, varname, embodiment)) == NULL)
		return FALSE;

	expr = var->Expression;

	return TRUE;
}

BOOL Document::VariableExprSet( ksAPI7::IVariable7Ptr & var, LPCWSTR expr )
{
	ksAPI7::IApplicationPtr app = var->Application;
	if (app == NULL)
		return FALSE;
	ksAPI7::IKompasErrorPtr errptr( app->KompasError );
	if (errptr == NULL)
		return FALSE;

	errptr->Clear();
	var->Expression = expr;
	if (errptr->Code != etSuccess) {
		errptr->Clear();
		return FALSE;
	}

	return TRUE;
}

BOOL Document::VariableExprSet( ksAPI7::IKompasDocument3DPtr & doc3d, LPCWSTR varname, LPCWSTR expr, LONG embodiment )
{
	ksAPI7::IVariable7Ptr var = NULL;
	if ((var = EmbodimentVariable( doc3d, varname, embodiment )) == NULL)
		return FALSE;

	return VariableExprSet( var, expr );
}

BOOL Document::CommentGet( ksAPI7::IKompasDocument3DPtr & doc3d, LPCWSTR & comment )
{
	// document 3D API7 to API5
	IDocument3DPtr doc3d5( IUnknownPtr(ksTransferInterface( doc3d, ksAPI3DCom, 0/*any document*/ ), false/*don't AddRef*/) );
	if (doc3d5 == NULL)
		return FALSE;

	comment = doc3d5->GetComment();
	// return empty string as no comment
	if (comment == NULL)
		comment = L"";

	return TRUE;
}

BOOL Document::CommentSet( ksAPI7::IKompasDocument3DPtr & doc3d, LPWSTR comment )
{
	// document 3D API7 to API5
	IDocument3DPtr doc3d5( IUnknownPtr(ksTransferInterface( doc3d, ksAPI3DCom, 0/*any document*/ ), false/*don't AddRef*/) );
	if (doc3d5 == NULL)
		return FALSE;

	// treat empty string as NULL to remove comment
	if ((comment != NULL) && wcslen(comment) == 0)
		comment = NULL;
	// Do not do update if values are equal
	LPWSTR curcom = doc3d5->GetComment();
	if (curcom == NULL) {
		if (comment == NULL)
			return TRUE;
	} else if ((comment != NULL) && (wcscmp(curcom, comment) == 0))
		return TRUE;

	BOOL ret = doc3d5->SetComment(comment);
	if (ret)
		ret = doc3d5->UpdateDocumentParam();

	return ret;
}

BOOL Document::AuthorGet( ksAPI7::IKompasDocument3DPtr & doc3d, LPCWSTR & author )
{
	// document 3D API7 to API5
	IDocument3DPtr doc3d5( IUnknownPtr(ksTransferInterface( doc3d, ksAPI3DCom, 0/*any document*/ ), false/*don't AddRef*/) );
	if (doc3d5 == NULL)
		return FALSE;

	author = doc3d5->GetAuthor();
	// return empty string as no author
	if (author == NULL)
		author = L"";
	return TRUE;
}

BOOL Document::AuthorSet( ksAPI7::IKompasDocument3DPtr & doc3d, LPWSTR author )
{
	// document 3D API7 to API5
	IDocument3DPtr doc3d5( IUnknownPtr(ksTransferInterface( doc3d, ksAPI3DCom, 0/*any document*/ ), false/*don't AddRef*/) );
	if (doc3d5 == NULL)
		return FALSE;

	// treat empty string as NULL to remove author
	if ((author != NULL) && wcslen(author) == 0)
		author = NULL;
	// Do not do update if values are equal
	LPWSTR curaut = doc3d5->GetAuthor();
	if (curaut == NULL) {
		if (author == NULL)
			return TRUE;
	} else if ((author != NULL) && (wcscmp(curaut, author) == 0))
		return TRUE;

	BOOL ret = doc3d5->SetAuthor(author);
	if (ret)
		ret = doc3d5->UpdateDocumentParam();

	return ret;
}

BOOL Document::NFeatureNameGet( ksAPI7::IKompasDocument3DPtr & doc3d, ULONG N, CString * name, BOOL append, LONG embodiment )
{
	ksAPI7::IEmbodimentsManagerPtr em(doc3d);
	if (em == NULL)
		return FALSE;
	ksAPI7::IEmbodimentPtr e = em->CurrentEmbodiment;
	if (e == NULL)
		return FALSE;
	ksAPI7::IPart7Ptr epart = e->GetPart();
	if (epart == NULL)
		return FALSE;
	ksAPI7::IFeature7Ptr feature(epart);
	if (feature == NULL)
		return FALSE;
	_variant_t subf(feature->GetSubFeatures(ksOperTree, TRUE, FALSE));
	if ((V_VT(&subf) != (VT_ARRAY | VT_DISPATCH)) || (V_ARRAY(&subf)->cDims != 1)) 
		return FALSE;
	ULONG count = V_ARRAY(&subf)->rgsabound[0].cElements - V_ARRAY(&subf)->rgsabound[0].lLbound;
	if (!(N < count))
		return FALSE;

	LPDISPATCH HUGEP * pSubf = NULL;
	BOOL ret = FALSE;
	HRESULT hr = ::SafeArrayAccessData( V_ARRAY(&subf), (void HUGEP* FAR*)&pSubf );
	if (!FAILED(hr) && pSubf) {
		ksAPI7::IFeature7Ptr subfptr(pSubf[N]);
		if (subfptr != NULL) {
			if (append)
				name->Append(subfptr->Name);
			else
				name->SetString(subfptr->Name);
		}

		ret = TRUE;
	}

	if ( pSubf )
		::SafeArrayUnaccessData( V_ARRAY(&subf) );

	return ret;
}

BOOL Document::FeatureNameEnum( ksAPI7::IKompasDocument3DPtr & doc3d, FEATURENAMES & names, LONG embodiment )
{
	ksAPI7::IEmbodimentsManagerPtr em(doc3d);
	if (em == NULL)
		return FALSE;
	ksAPI7::IEmbodimentPtr e = em->CurrentEmbodiment;
	if (e == NULL)
		return FALSE;
	ksAPI7::IPart7Ptr epart = e->GetPart();
	if (epart == NULL)
		return FALSE;
	ksAPI7::IFeature7Ptr feature(epart);
	if (feature == NULL)
		return FALSE;
	_variant_t subf(feature->GetSubFeatures(ksOperTree, TRUE, FALSE));
	if ((V_VT(&subf) != (VT_ARRAY | VT_DISPATCH)) || (V_ARRAY(&subf)->cDims != 1)) 
		return FALSE;
	ULONG count = V_ARRAY(&subf)->rgsabound[0].cElements - V_ARRAY(&subf)->rgsabound[0].lLbound;

	LPDISPATCH HUGEP * pSubf = NULL;
	HRESULT hr = ::SafeArrayAccessData( V_ARRAY(&subf), (void HUGEP* FAR*)&pSubf );
	BOOL ret = TRUE;
	if (!FAILED(hr) && pSubf) {
		names.clear();
		for (ULONG index = 0; index < count; index++) {
			ksAPI7::IFeature7Ptr subfptr(pSubf[index]);
			if (subfptr != NULL) {
				FEATURENAME name;
				names.push_back(name);
				FEATURENAME *pname = &names.back();
				pname->name = (LPCWSTR)subfptr->Name;
				pname->excluded = subfptr->Excluded;
				ksAPI7::IModelObjectPtr suboptr(subfptr);
				if (suboptr != NULL) {
					pname->type = suboptr->ModelObjectType;
					pname->hidden = suboptr->Hidden;
				} else {
					pname->type = o3d_unknown;
					pname->hidden = TRUE;
				}
			} else {
				ret = FALSE;
				break;
			}
		}
	}

	if ( pSubf )
		::SafeArrayUnaccessData( V_ARRAY(&subf) );

	return ret;
}
