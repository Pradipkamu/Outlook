@echo off
if not exist "D:\FollowupOrganizer\data" mkdir "D:\FollowupOrganizer\data"
>"D:\FollowupOrganizer\data\STOP" echo User requested emergency stop. Review before resuming.
echo Automatic dispatch blocked. Already submitted emails cannot be recalled.
pause
