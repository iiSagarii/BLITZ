## BLITZ user-guide ##

--> Set-key: 
	1] If the API key has already been set and is not expired, continue on to the "TSS validation" function.
	2] If using BLITZ for the first time OR the API key has been expired, continue with the "Set-key" function.
		o Paste the key in the space provided and click on the "OK" button.
		o The key would be set and BLITZ would be ready to use. 
		

--> TSS validation:
	1] Select the "TOE Type" first. (If the TOE Type has not been selected, the rest of the functions would be disable and the process would not kickstart):
		o This would consists of 2 options: "DISTRIBUTED" and "STANDALONE".
		o Select the TOE type carefully and accurately as BLITZ's processing for both the types would differ significantly.
		o Once the TOE type has been selected, click on the "Proceed" button; this would kickstart the processing element.
	
	2] Select the ST:
		o Click on the "Select ST" button and browse for the latest version of the ST document for that particular project.
		o The ST should be present on the machine that BLITZ is being operated on. The selected, would be considered for further processing.
	
	3] Claims conformance with:
		o This section is divided into 3 sub-sections. Click on the checkbox and select the conformance profiles that the ST conforms with (e.g.: NDcPP_v3.0, MOD_IPS_v1.0, etc).
		o Select all the applicable PPs, MODs and PKGS for the project as per the ST.
	
	4] Validate TSS:
		o Once all the steps above have been followed accurately, BLITZ would start the validation for the TSS assurance activities for the selected ST.


--> The output would be stored in the same directory as the code is being run, under a folder named "BLITZ-output".