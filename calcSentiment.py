from textblob import TextBlob 

#Calculates the polarity of the review title 
def calcTPolarity(calcTPol):

    sentBlob = TextBlob(calcTPol)
    sentPol = sentBlob.sentiment.polarity
    #returns the calculation
    return sentPol

#Calculats the subjectivity of the review title 
def calcTSubjectivity(calcTSub):

    sentBlob = TextBlob(calcTSub)
    sentSub = sentBlob.sentiment.subjectivity
    #returns the calculation
    return sentSub

#Calculates the polarity of the review
def calcRPolarity(calcRPol):

    sentBlob = TextBlob(calcRPol)
    sentPol = sentBlob.sentiment.polarity
    #returns the calculation
    return sentPol

#Calculates the subjectivity of the review   
def calcRSubjectivity(calcRSub):

    sentBlob = TextBlob(calcRSub)
    sentPol = sentBlob.sentiment.subjectivity
    #returns the calculation
    return sentPol
    
