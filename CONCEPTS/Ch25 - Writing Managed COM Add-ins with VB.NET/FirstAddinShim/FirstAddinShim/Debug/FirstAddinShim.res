        ��  ��                  �      �� ��     0 	        �4   V S _ V E R S I O N _ I N F O     ���               ?                        <   S t r i n g F i l e I n f o      0 4 0 9 0 4 B 0   8   C o m p a n y N a m e     C o m p a n y N a m e   H   F i l e D e s c r i p t i o n     F i l e D e s c r i p t i o n   0   F i l e V e r s i o n     1 . 0 . 0 . 0   D   L e g a l C o p y r i g h t   ( c )   C o m p a n y N a m e   F   I n t e r n a l N a m e   F i r s t A d d i n S h i m . d l l     N   O r i g i n a l F i l e n a m e   F i r s t A d d i n S h i m . d l l     >   P r o d u c t N a m e     F i r s t A d d i n S h i m     4   P r o d u c t V e r s i o n   1 . 0 . 0 . 0   D    V a r F i l e I n f o     $    T r a n s l a t i o n     	�  0   R E G I S T R Y   ��f       0 	        HKCR
{
	FirstAddin.Connect = s 'Connect Class'
	{
		CLSID = s '{b7b5d1ed-8528-4c11-b70a-ee1bd7e4d1bd}'
	}
	NoRemove CLSID
	{
		ForceRemove '{b7b5d1ed-8528-4c11-b70a-ee1bd7e4d1bd}' = s 'FirstAddin.Connect'
		{
			ProgID = s 'FirstAddin.Connect'
			InprocServer32 = s '%MODULE%'
			{
				val ThreadingModel = s 'Apartment'
			}
			
		}
	}
}

HKCU
{
	NoRemove Software
	{
		NoRemove Microsoft
		{
			NoRemove Office
			{
				NoRemove Excel
				{
					NoRemove Addins
					{
						ForceRemove FirstAddin.Connect
						{
							val 'Description' = s 'Chapter 25 - Writing Managed COM Add-ins with VB.NET'
							val 'FriendlyName' = s 'First Shimmed Add-in'
							val 'LoadBehavior' = d 3
						}
					}
				}
			}
		}
	}
}





  <       �� ��     0	                 F i r s t A d d i n S h i m                       