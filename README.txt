The purpose of this script is to prefill a FedRAMP Risk Exposure Table (RET) to the greatest extent possible through automation from a completed Security Requirements Traceability Matrix (SRTM)

Updated to follow rev5 tempaltes provided by FedRAMP

Assumptions:
-SRTM IS COMPLETE AND READY TO DELIVER; no more findings, all custom formating removed, all checks have been run and reviewed

-If any changes are made to the SRTM, all changes must also be updated in the RET, else the RET process/script restarted from scratch.

-Indentified risks, risk statements, and reccomendations should all be written as a single paragraph in the SRTM. 
---Multiple sentances can be used, but newlines are used to separate multiple findings in a single Assessment Procedure. 
---Failure to folow this will result in ValueErrors and that control family failing to be processed
---Same rule applies to PL-2 Findings (SSP Implementation Differential?)

-For any "refer to" or "See" Findings - these are restricted to referancing another finding within the control or enhancement, not outside of it.
-A similar restriction exists for PL-2 Findings. At least 1 PL-2 needs to have the direct statement instead of referancing antoher PL-2 or regular finding.
--"Findings/PL-2s that begin with 'Refer to' or "See" will not be processed and added to the RET" 
--"Refer to" findings should be limited to the same type of finding wherever possible (regular findings referance regular findings, PL-2s referance PL-2s)


-I have attempted to account for the most reasonable typos and variations, but the RET should still be given a onceover to ensure that all data transfered correctly.
