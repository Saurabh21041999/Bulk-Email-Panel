#!/usr/bin/env python
# coding: utf-8

# In[2]:


import smtplib                                          #use to deal with mails of gmail
def email_send_fun(to,subject,message,from_,pass_):
    
    s=smtplib.SMTP("smtp.gmail.com",587)                #create session for gmail
    s.starttls()                                       #activate transport layer
    s.login(from_,pass_)
    msg="subject: {}\n\n{}".format(subject,message)
    s.sendmail(from_,to,msg)
    x=s.ehlo()                                            #Return status
    if x[0]==250:
        return "s"
    else:
        return "f"
    s.close()
   


# In[ ]:





# In[ ]:




