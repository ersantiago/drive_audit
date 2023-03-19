#====================================================================#
#                     Files directory Audit                          #
#====================================================================#

Utility to scan a directory recursively and check inside the files for any flagged terms from a list
  e.g. Finding any file that contains a word matching "emerson" inside a directory of over 1000 files.

1. Supported formats: xlsx, xls, docx, doc, pdf, csv, odf, odt & other plain-text formats (txt, log, ini, cfg, sh, bat)
2. Output format: fullpath,filename,owner,filetype,timestamp,flag_chk,flagged_terms
3. Runs recursively
