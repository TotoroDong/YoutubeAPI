# -*- coding: utf-8 -*-

import os
import pandas as pd


import google_auth_oauthlib.flow
import googleapiclient.discovery
import googleapiclient.errors

scopes = ["https://www.googleapis.com/auth/youtube.readonly"]

def main():
    # Disable OAuthlib's HTTPS verification when running locally.
    # *DO NOT* leave this option enabled in production.
    os.environ["OAUTHLIB_INSECURE_TRANSPORT"] = "1"

    api_service_name = "youtube"
    api_version = "v3"
    client_secrets_file = ".json" #Replace this with your json file name

    # Get credentials and create an API client
    flow = google_auth_oauthlib.flow.InstalledAppFlow.from_client_secrets_file(
        client_secrets_file, scopes)
    credentials = flow.run_local_server(port=0)
    youtube = googleapiclient.discovery.build(
        api_service_name, api_version, credentials=credentials)


    request = youtube.search().list(
        q='university student mental health',
        type="video",
        part="id,snippet",
        maxResults=100,
        publishedAfter="2022-08-1T00:00:00Z",
        publishedBefore="2023-07-19T00:00:00Z"
    )
    search_response = request.execute()
    
    videos = []  
    for search_result in search_response['items']:
        try:
            video_id = search_result['id']['videoId']
            title = search_result['snippet']['title']
            description = search_result['snippet']['description']

            # Get statistics for the video
            video_request = youtube.videos().list(
                part="snippet,statistics",
                id=video_id
            )
            video_response = video_request.execute()

            # Get the number of likes, dislikes, views, and comment count if available
            likes = video_response['items'][0]['statistics'].get('likeCount', None)
            views = video_response['items'][0]['statistics'].get('viewCount', None)
            comment_count = video_response['items'][0]['statistics'].get('commentCount', None)

            # Get tags (keywords), if available
            tags = video_response['items'][0]['snippet'].get('tags') if 'snippet' in video_response['items'][0] else None

            # Get channel ID from the search result
            channel_id = search_result['snippet']['channelId']
            
            # Make a request to get information about the channel
            channel_request = youtube.channels().list(
                part="statistics",
                id=channel_id
            )
            channel_response = channel_request.execute()

            # Get the subscriber count for the channel
            subscriber_count = channel_response['items'][0]['statistics'].get('subscriberCount', None)
        
            # Check if subscriber count is greater than 20000 before appending video
            if subscriber_count is not None and int(subscriber_count) > 20000:
                videos.append([video_id, title, description, likes, views, subscriber_count, tags, comment_count])
        except Exception as e:
            print(f"An error occurred for video {video_id}: {e}")

    df = pd.DataFrame(videos, columns=["video_id", "title", "description", "likes", "views", "subscriber_count", "tags", "comment_count"])

    # Export DataFrame to Excel
    df.to_excel("youtube_search_results.xlsx", index=False)

if __name__ == "__main__":
    main()
