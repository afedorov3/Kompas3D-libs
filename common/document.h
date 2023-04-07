#ifndef _DOCUMENT_H
#define _DOCUMENT_H

namespace Document {
	enum {
		EMBODIMENT_CURRENT = -1,
		EMBODIMENT_TOP     =  0
	};

	typedef struct {
		CString name;
		CString number;
	} EMBODIMENTNAME;
	typedef std::vector <EMBODIMENTNAME> EMBODIMENTNAMES;

	typedef struct {
		CString name;
		ksObj3dTypeEnum type;
		BOOL excluded;
		BOOL hidden;
	} FEATURENAME;
	typedef std::vector <FEATURENAME> FEATURENAMES;

	// document manage
	BOOL IsDocumentValid( ksAPI7::IKompasDocument3DPtr & doc3d );
	BOOL RebuildDocument( ksAPI7::IKompasDocument3DPtr & doc3d );
	BOOL RebuildModelEx( ksAPI7::IKompasDocument3DPtr & doc3d, BOOL redraw = TRUE );

	// embodiment
	LONG EmbodimentCount( ksAPI7::IKompasDocument3DPtr & doc3d );
	ksAPI7::IEmbodimentPtr EmbodimentPtrGet( ksAPI7::IKompasDocument3DPtr & doc3d, LONG index = EMBODIMENT_TOP );
	BOOL EmbodimentGet( ksAPI7::IKompasDocument3DPtr & doc3d, LONG & index );
	BOOL EmbodimentSet( ksAPI7::IKompasDocument3DPtr & doc3d, LONG index );
	BOOL EmbodimentMarkingGet( ksAPI7::IKompasDocument3DPtr & doc3d, LONG index, LONG type, CString * name, BOOL append = FALSE );
	BOOL EmbodimentNameEnum( ksAPI7::IKompasDocument3DPtr & doc3d, EMBODIMENTNAMES & names );

	// name
	BOOL ModelEmbodimentNameGet( ksAPI7::IKompasDocument3DPtr & doc3d, LONG embodiment, LPWSTR & modelname );
	BOOL ModelEmbodimentNameSet( ksAPI7::IKompasDocument3DPtr & doc3d, LONG embodiment, LPCWSTR modelname );
	BOOL ModelEmbodimentNameSetFmt( ksAPI7::IKompasDocument3DPtr & doc3d, LONG embodiment, LPCWSTR fmt, ... );
	BOOL ModelNameGet( ksAPI7::IKompasDocument3DPtr & doc3d, LPWSTR & modelname );
	BOOL ModelNameSet( ksAPI7::IKompasDocument3DPtr & doc3d, LPCWSTR modelname );
	BOOL ModelNameSetFmt( ksAPI7::IKompasDocument3DPtr & doc3d, LPCWSTR fmt, ... );

	// variables
	BOOL IsVariableNameCharValid( WCHAR ch, BOOL first = FALSE );
	BOOL IsVariableNameValid( LPCWSTR name, LONG len = -1 );
	ksAPI7::IVariable7Ptr EmbodimentVariable( ksAPI7::IKompasDocument3DPtr & doc3d, LPCWSTR varname, LONG embodiment = EMBODIMENT_TOP, BOOL create = FALSE );
	BOOL VariableValueGet( ksAPI7::IKompasDocument3DPtr & doc3d, LPCWSTR varname, DOUBLE & value, LONG embodiment = EMBODIMENT_TOP );
	BOOL VariableValueSet( ksAPI7::IKompasDocument3DPtr & doc3d, LPCWSTR varname, DOUBLE value, LONG embodiment = EMBODIMENT_TOP );
	BOOL VariableExprGet( ksAPI7::IKompasDocument3DPtr & doc3d, LPCWSTR varname, LPCWSTR & expr, LONG embodiment = EMBODIMENT_TOP );
	BOOL VariableExprSet( ksAPI7::IVariable7Ptr & var, LPCWSTR expr );
	BOOL VariableExprSet( ksAPI7::IKompasDocument3DPtr & doc3d, LPCWSTR varname, LPCWSTR expr, LONG embodiment = EMBODIMENT_TOP );

	// comment
	BOOL CommentGet( ksAPI7::IKompasDocument3DPtr & doc3d, LPCWSTR & comment );
	BOOL CommentSet( ksAPI7::IKompasDocument3DPtr & doc3d, LPWSTR comment );

	// author
	BOOL AuthorGet( ksAPI7::IKompasDocument3DPtr & doc3d, LPCWSTR & author );
	BOOL AuthorSet( ksAPI7::IKompasDocument3DPtr & doc3d, LPWSTR author );

	// feature name
	BOOL NFeatureNameGet( ksAPI7::IKompasDocument3DPtr & doc3d, ULONG N, CString * name, BOOL append = FALSE, LONG embodiment = EMBODIMENT_TOP );
	BOOL FeatureNameEnum( ksAPI7::IKompasDocument3DPtr & doc3d, FEATURENAMES & names, LONG embodiment = EMBODIMENT_TOP );

} /* Document */

#endif /* _DOCUMENT_H */
