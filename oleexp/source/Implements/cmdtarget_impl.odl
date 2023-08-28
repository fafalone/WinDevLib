[
    odl,
    uuid(b722bccb-4e68-101b-a2bc-00aa00404770),
]
interface IOleCommandTarget : stdole.IUnknown {

    HRESULT QueryStatus(
        [in, out] UUID *pguidCmdGroup,
        [in] LONG cCmds,
        [in, out] OLECMD *prgCmds,
        [in, out] OLECMDTEXT *pCmdText);

    HRESULT Exec(
        [in, out] UUID *pguidCmdGroup,
        [in] LONG nCmdID,
        [in] OLECMDEXECOPT nCmdexecopt,
        [in] VARIANT *pvaIn,
        [in, out] VARIANT *pvaOut);
}
