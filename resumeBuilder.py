from groq import Groq
import os
import datetime


client = Groq(
    api_key="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",
)

import pdfplumber

resume = ""
with pdfplumber.open("Hitesh Soneta.pdf") as pdf:
    for page in pdf.pages:
        resume += page.extract_text() + "\n"  # Extract text




with open('jobDescription.txt', 'r') as job_description_file:
    jobDescription = job_description_file.read()

completion = client.chat.completions.create(
    model="llama-3.3-70b-versatile",
    messages=[
        {
            "role": "user",
            "content": "Build a custom resume for this job posting, just modify experience part thats it here is the resume:" + resume + "  and here is the job description " + jobDescription
        },
        {
            "role": "assistant",
            "content": "Please provide the job posting details, and I'll create a custom resume tailored to the job requirements.\n\nPlease provide the following information:\n\n1. Job title\n2. Job description\n3. Requirements (e.g., skills, experience, education)\n4. Any specific keywords or phrases mentioned in the job posting\n\nOnce I have this information, I'll create a custom resume that highlights your relevant skills and experiences, increasing your chances of getting noticed by the hiring manager."
        }
    ],
    temperature=1,
    max_tokens=1024,
    top_p=1,
    stream=False,
    stop=None,
)


timestamp = datetime.datetime.now().strftime("%Y%m%d%H%M%S")
output_file_name = f"resume_{timestamp}.md"    
with open(output_file_name, 'w') as output_file:
    output_file.write(completion.choices[0].message.content)