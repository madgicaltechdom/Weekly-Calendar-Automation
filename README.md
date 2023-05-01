# Weekly-Calendar-Automation
Sync the calendar with google sheet and validate the weekly tasks

# Create Pivot Table
1. Select the sheet where you would like to create a pivot table
2. Open script editor and run 'TaskPivot' in macros.gs

# Send Mail Nodification
This feature adds to createTrigger2() function which runs every day. It is check pivot table condition is followed properly or not. if it does not properly follow pivot table condition then one mail is sent to the user with alert mail and conditions which are not properly followed by the user.

# Apply Trigger 
Our sheet user need to apply Trigger in his Sheet so the sheet can run automatically by the help of trigger function.
1. Go on Extensions > Go on Apps Script
![image](https://user-images.githubusercontent.com/109965832/235437861-f9307965-0c5f-48dd-a095-0c0ab7de8dd1.png)

2. After that Apps Script is open so go on applyTrigger.gs.

![image](https://user-images.githubusercontent.com/109965832/235438425-14b4eac6-b70d-4c96-8ad8-9e6faf0d5747.png)

3. So now you can see two functions :-
      * createTrigger1()
      * createTrigger2()
      
![image](https://user-images.githubusercontent.com/109965832/235438704-30c5c278-73a1-42bf-9fe7-648e7ff5b9dd.png)

4. Then run this file with both functions one by one. Clicking on the run button.
![image](https://user-images.githubusercontent.com/109965832/235439094-f99c0cd0-638c-4adb-8729-ee6514cb0ec6.png)

5. If the User wants to see trigger is applied or not so go on triggers.

![image](https://user-images.githubusercontent.com/109965832/235440110-5280ec31-25c5-4049-a700-71175931824a.png)
