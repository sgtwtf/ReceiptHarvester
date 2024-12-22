Please note that while Albertsons, Haggen, and Safeway etc.. appear to share the same backend, it’s unclear whether their websites behave in the same way. 
When running this script for the first time, I recommend adding some longer sleep durations. 
This will give you enough time to adjust the click positions if necessary.

Use Pulover's Macro Creator (https://www.macrocreator.com/download/) and save the safeway.pmc to your device.

Login to your account.
go to wallet then click on Digital receipts.
Make this page the left half of your screen.

Update the macro postions to match your resolution.

(
Line 2: This line is used to handle scenarios where the script fails and you need to restart. You input the number of tabs to begin with, ensuring that the process resumes from the correct point.
  You only need to change this if your script fails for some reason.
  
Line 4: Update this line with the window currently displayed on the left half of your screen, specifically the one showing the list of receipts.
Line 7: Perform the save operation here and also press the "4" key (this step may not be strictly necessary, but it has helped ensure better results in my experience).
Line 8: This marks the start of the logic loop. There are two key variables that determine its behavior:
  clicking_to_be_done: This variable tracks the total number of tabs that need to be completed. It's essential because without it, the script wouldn't know how many tabs remain after each round.
  clicking_done: This variable counts how many tabs have been completed in the current round. If fewer tabs are completed than required, the script will continue to the next tab until the goal is met.
Line 12: Resets the clicking_to_be_done counter to zero, ensuring that the next round starts fresh from the top.
Line 13: This line looks for the border that indicates which date the tab corresponds to, and when the required number of tabs for the current round is reached, it moves the cursor to the correct position.
Line 15: Moves the cursor slightly from the border and clicks on the receipt to open it.
Line 17: Clicks on the envelope icon (indicating the email option).
Line 18: Focuses the cursor on the input box, where you can enter any email address.
Line 19: Types the email address where you want the receipt to be sent.
Line 20: Clicks the "Send" button to send the email.
Line 21: Waits for 1 second to allow the email to be sent successfully.
Line 22: Clicks the "Return" button to close the current receipt view.
Line 23: Pauses for 1 second before proceeding to the next action.
Line 24: Loops back to Line 3 to repeat the process, ensuring the next tab is processed.
)

When the last receipt is processed, the script encounters an issue where the border box cannot be found, causing it to hang and preventing further execution.
