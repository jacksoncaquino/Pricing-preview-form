# Pricing Preview Form
Designed a form using MS-Forms so recruiters and HR Business partners who are considering activating jobs in new locations can request to know pay ranges without going through the whole 2-weeks long process to get the job created and priced in our HR information system. Automated the notification to the Compensation point of contact to know of the new job that needed pricing. Automated the feedback email for the compensation point of contact to reply with the results of the pricing. The process came from 2 weeks to max 2 business days. Used Typescript, Power Automate, and VBA for Excel.


# The form
This is what the intake form looks like:
![image](https://github.com/jacksoncaquino/Pricing-preview-form/assets/61064363/474014c3-01c2-431c-8233-98c1f4530063)

The user can choose to input either job family and level or job code. If they choose job family and level, there is a list with job families they can choose from and another list where they can choose the level. If they choose to provide job code, they just need to type the job code and submit the information.

# What happens when there's a new submission on the form?
I have created a flow using Power Automate, which is a low code solution. The flow is triggered by the form submission and then it takes the following steps:<br>
• Gets the details from the form entry so we can retrieve it on the later steps of the process<br>
• Gets the information about the requestor so we can know their name, job title, and department from their Microsoft Office profile<br>
• Gets the job code from the form entry<br>
• Checks if the job code is blank so it knows it's a junction of job family and level<br>
• If it's job family and level, it concatenates job family/level<br>
• Initiates a variable with the requestors first and last name to add to the Excel tracker<br>
• With Job Code (or the equivalent job family/level combination), compensation market, and name of the requestor, it runs a TypeScript that loads these fields as a new row to our pricing tracker (see the code of the script attached)<br>
• Gets the name of the compensation professional who is assigned to price the new job<br>
• Sends the email to the compensation professional who will price the job, along with instructions on how to send the feedback to the requestor<br>
![image](https://github.com/jacksoncaquino/Pricing-preview-form/assets/61064363/04521024-8deb-45cd-bca1-075afb2013f8)

# How does the compensation team price the jobs?
Currently it takes less than two minutes using the midpoint prediction tool. Refer to the project here to know more about it: [https://github.com/jacksoncaquino/Midpoint-Prediction-Tool/](https://github.com/jacksoncaquino/Midpoint-Prediction-Tool/)

# When the job is priced, how to send the feedback to the requestor?
On the Excel add-in I created and shared with the team there is a button that allows the compensation team to send the response directly to the requestors:
![image](https://github.com/jacksoncaquino/Pricing-preview-form/assets/61064363/40c18f9f-71be-4976-8480-653ba6471c80)

VBA writes an email with the information and sends it on behalf of the compensation point of contact. The email looks like this:
![image](https://github.com/jacksoncaquino/Pricing-preview-form/assets/61064363/a7a15c10-4247-4355-9b0d-1f29f9bca534)

