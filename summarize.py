"""Text Summary Generator

This script will take in an argument (URL Only currently), extract the text,
parse it through the spacy NLP then generate a reduced summary of the text.
"""

import argparse
import os
import win32com.client
from newspaper import Article
import spacy
from spacy.lang.en.stop_words import STOP_WORDS
from string import punctuation
from heapq import nlargest
from datetime import datetime, timedelta


def summarize(text, per):
    nlp = spacy.load('en_core_web_sm')
    doc = nlp(text)
    tokens=[token.text for token in doc]
    word_frequencies={}
    for word in doc:
        if word.text.lower() not in list(STOP_WORDS):
            if word.text.lower() not in punctuation:
                if word.text not in word_frequencies.keys():
                    word_frequencies[word.text] = 1
                else:
                    word_frequencies[word.text] += 1

    max_frequency = max(word_frequencies.values())
    for word in word_frequencies.keys():
        word_frequencies[word]=word_frequencies[word]/max_frequency
    
    sentence_tokens = [sent for sent in doc.sents]
    sentence_scores = {}
    for sent in sentence_tokens:
        for word in sent:
            if word.text.lower() in word_frequencies.keys():
                if sent not in sentence_scores.keys():
                    if sent not in sentence_scores.keys():
                        sentence_scores[sent]=word_frequencies[word.text.lower()]
                    else:
                        sentence_scores[sent]+=word_frequencies[word.text.lower()]
    
    select_length=int(len(sentence_tokens)*per)
    summary = nlargest(select_length, sentence_scores, key=sentence_scores.get)
    final_summary = [word.text for word in summary]
    summary=' '.join(final_summary)
    return summary


def verbose_msg(msgtxt, verbose=False):
    '''Checks if verbose mode is active and prints a msg to command line if it is'''
    if verbose:
        print(msgtxt)


def summarize_url(url, per=0.1):

    article = Article(url)
    article.download()
    article.parse()
    
    return summarize(article.text, per)

def __main__():

    parser = argparse.ArgumentParser()
    parser.add_argument("--url", "-u", help="Pulls text from provided URL and returns summary")
    parser.add_argument("--size", "-s", help="Set the size of the summary as a percentage of the original text. Takes a decimal from 0 - 1. Default is 0.1")
    parser.add_argument("--verbose", help="Will print status messages has script is running. Useful for debugging an issue.", action="store_true")
    parser.add_argument("--email", help="Flag to scan Email Inbox and provide summary of messages", action="store_true")
    args = parser.parse_args()
    
    if args.size: #Set the Summary size

        summary_rate = float(args.size)

    else: # Default to 10% if no rate is provided
        summary_rate = 0.1

    parser_text = []
    if args.url:

        verbose_msg(f"Fetching {args.url}", args.verbose)
        article = Article(args.url)
        verbose_msg("Downloading Content from URL", args.verbose)
        article.download()
        verbose_msg("Parsing Text", args.verbose)
        article.parse()
        parser_text.append(article.text)
    
    elif args.email:
        verbose_msg("This is where I would Parse my E-mail", args.verbose)
        outlook = win32com.client.Dispatch('outlook.application')
        mapi = outlook.GetNamespace("MAPI")

        inbox = mapi.getDefaultFolder(6)
        messages = inbox.Items

        recieved_dt = datetime.now() - timedelta(days=1)
        recieved_dt = recieved_dt.date().strftime('%m/%d/%Y')
        messages = messages.restrict("[ReceivedTime] >= '" + recieved_dt + "'")

        for message in messages:
            parser_text.append(message.body)


    verbose_msg("Generating Summary", args.verbose)
    for text in parser_text:
        print(summarize(text, summary_rate))
        if len(parser_text) > 1:
            print("\n"+("-" * 10)+ "\n")


if __name__ == "__main__":
    __main__()
