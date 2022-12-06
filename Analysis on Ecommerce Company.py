#!/usr/bin/env python
# coding: utf-8

# Importing Libraries

# In[1]:


import pandas as pd
import numpy as np
import seaborn as sd


# Importing Files

# In[2]:


order=pd.read_excel('Company X - Order Report.xlsx')
pincode=pd.read_excel('Company X - Pincode Zones.xlsx')
sku=pd.read_excel('Company X - SKU Master.xlsx')
invoice=pd.read_excel('Courier Company - Invoice.xlsx')
rates=pd.read_excel('Courier Company - Rates.xlsx')


# In[3]:


df= pd.merge(order,sku)
df2 = invoice


# In[4]:


df.rename(columns = {'ExternOrderNo':'Order ID'}, inplace = True)
df.head()


# Calculating Total Weight

# In[5]:


total_weight= df['Weight (g)'] * df['Order Qty']
df['total_weight']=total_weight
df.head()


# In[6]:


df=df.groupby(['Order ID']).sum()
df.head(125)


# In[7]:


df.reset_index('Order ID',inplace=True)
df.head()


# Calculating Total Weight (kG):

# In[8]:


total_weigh_kg=df['total_weight']/1000
df['total_weigh_kg']=total_weigh_kg
df.head()


# creating Weight slab as per X (KG)

# In[9]:


Weight_slab=[]
for l in df['total_weigh_kg']:
  if l<=0.5:
    Weight_slab.append(0.5)
  elif l>=0.5 and l<=1:
    Weight_slab.append(1.0)
  elif l>=1 and l<=1.5:
    Weight_slab.append(1.5)
  elif l>=1.5 and l<=2:
    Weight_slab.append(2.0)
  elif l>=2 and l<=2.5:
    Weight_slab.append(2.5)
  elif l>=2.5 and l<=3:
    Weight_slab.append(3.0)
  elif l>=3 and l<=3.5:
    Weight_slab.append(3.5)
  elif l>=3.5 and l<=4:
    Weight_slab.append(4)
  else:
    Weight_slab.append(4.5)


# In[10]:


df['Weight_slab']=Weight_slab
df.head()


# In[11]:


pincode.rename(columns = {'Customer Pincode':'Customer_Pincode(X)'}, inplace = True)
pincode.head()


# In[12]:


pincode.rename(columns = {'Zone':'Zone(X)'}, inplace = True)
pincode.head()


# In[13]:


df2=pd.concat([df2,pincode],axis=1)
df2.head()


# In[14]:


df2.shape


# Total weight as per Courier Company (KG)

# In[15]:


total_weight_as_per_CC=df2['Charged Weight']
df2['total_weight_as_per_CC']=total_weight_as_per_CC
df2.head()


# Weight slab charged by Courier Company (KG)

# In[16]:


Weight_slab_CC=[]
for l in total_weight_as_per_CC:
  if l<=0.5:
    Weight_slab_CC.append(0.5)
  elif l>=0.5 and l<=1:
    Weight_slab_CC.append(1.0)
  elif l>=1 and l<=1.5:
    Weight_slab_CC.append(1.5)
  elif l>=1.5 and l<=2:
    Weight_slab_CC.append(2.0)
  elif l>=2 and l<=2.5:
    Weight_slab_CC.append(2.5)
  elif l>=2.5 and l<=3:
    Weight_slab_CC.append(3.0)
  elif l>=3 and l<=3.5:
    Weight_slab_CC.append(3.5)
  elif l>=3.5 and l<=4:
    Weight_slab_CC.append(4)
  else:
    Weight_slab_CC.append(4.5)


# In[17]:


df2['Weight_slab_CC']=Weight_slab_CC
df2.head()


# Cross Checking values Customer Pincode and Customer Pincode(X) is right or wrong. 

# In[18]:


df2['Customer Pincode'].equals(df2['Customer_Pincode(X)'])


# Delivery Zone as per X

# In[19]:


df3=df2[['Order ID','Customer_Pincode(X)','Zone(X)','Type of Shipment']]
df3


# In[20]:


p=df.merge(df3,on=['Order ID'])
p


# Calculating Billing Amount as per X

# In[21]:


total_weight=p['total_weigh_kg']
word =p['Zone(X)']
extracharges=p['Type of Shipment']
Billing_Amount_as_per_X=[]
def Charges(count1,i,fwd_base,fwd_additional,rto_base,rto_additional):    
    if extracharges[count1] == "Forward charges":         
            noC = 0
            sumW = 0
            sumF = 0
            tempW = int(i)
            tempF = i - tempW

            for j in range(0, tempW):
                sumW += fwd_base*2
                noC = noC+1

            if tempF != 0:
                if noC >= 1:
                    if tempF > 0.5:
                        sumF = fwd_base+fwd_additional

                    elif tempF == 0.5:
                        sumF = fwd_base
                    else:
                        sumF = fwd_additional
                    sum = sumW+sumF
                    Billing_Amount_as_per_X.append(sum)
                    print(i, sum, "Forward charges")

                else:
                    if tempF > 0.5:
                        sumF = fwd_base+fwd_additional

                    elif tempF == 0.5:
                        sumF = fwd_base

                    else:
                        sumF = fwd_base
                    Billing_Amount_as_per_X.append(sumF)
                    print(i, sumF, "Forward charges")
            else:
                sum = sumW+sumF
                Billing_Amount_as_per_X.append(sum)
                print(i, sum, "Forward charges")

    else:
        noC = 0
        sumW = 0
        sumF = 0
        tempW = int(i)
        tempF = i - tempW
        rtorateW = 0
        rtorateF = 0

        for j in range(0, tempW):
            sumW += fwd_base*2
            rtorateW += rto_base*2
            noC = noC+1

        if tempF != 0:
            if noC >= 1:
                if tempF > 0.5:
                    sumF = fwd_base+fwd_additional
                    rtorateF = rto_base+rto_additional

                elif tempF == 0.5:
                    sumF = fwd_base
                    rtorateF = rto_base
                else:
                    sumF = fwd_additional
                    rtorateF = rto_additional
                sum = sumW+sumF+rtorateF+rtorateW
                Billing_Amount_as_per_X.append(sum)
                print(i, sum, "Forward and RTO charges")

            else:
                if tempF > 0.5:
                    sumF = fwd_base+fwd_additional
                    rtorateF = rto_base+rto_additional

                elif tempF == 0.5:
                    sumF = fwd_base
                    rtorateF = rto_base

                else:
                    sumF = fwd_base
                    rtorateF = rto_base
                sum = sumF+rtorateF+rtorateW
                Billing_Amount_as_per_X.append(sum)
                print(i, sum, "Forward and RTO charges")
        else:
            sum = sumW+sumF+rtorateF+rtorateW
            Billing_Amount_as_per_X.append(sum)
            print(i, sum, "Forward and RTO charges")

for count1,i in enumerate(total_weight):
    
    if word[count1] == "b":
        fwd_base = 33
        fwd_additional = 28.3
        rto_base = 20.5
        rto_additional = 28.3
        Charges(count1,i,fwd_base,fwd_additional,rto_base,rto_additional)
        
    if word[count1] == "d":
        
        fwd_base = 45.4
        fwd_additional = 44.8
        rto_base = 41.3
        rto_additional = 44.8
        Charges(count1,i,fwd_base,fwd_additional,rto_base,rto_additional)
        
    if word[count1] == "e":
        fwd_base = 56.5
        fwd_additional =55.5
        rto_base = 50.7
        rto_additional = 55.5
        Charges(count1,i,fwd_base,fwd_additional,rto_base,rto_additional)


# In[22]:


len(Billing_Amount_as_per_X)


# In[23]:


p['Billing_Amount_as_per_X']=Billing_Amount_as_per_X
p.head(124)


# In[24]:


f=df2.merge(p,on=['Order ID'])
f.head()


# In[25]:


f.columns


# Arranging in Order

# In[26]:


result=f[['AWB Code','Order ID','total_weigh_kg', 'Weight_slab','total_weight_as_per_CC', 'Weight_slab_CC','Zone(X)_y','Zone','Billing_Amount_as_per_X','Billing Amount (Rs.)']]


# In[27]:


result


# In[28]:


Difference_Between_Expected_Charges_and_Billed_Charges=round(result['Billing_Amount_as_per_X']-result['Billing Amount (Rs.)']) 
result['Difference_Between_Expected_Charges_and_Billed_Charges']=Difference_Between_Expected_Charges_and_Billed_Charges


# In[29]:


result.head()


# Expoting Data into Excel

# In[30]:


result.to_excel('result1.xlsx',index=False)

