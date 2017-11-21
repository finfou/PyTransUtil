#!/usr/bin/env python
#! -*- coding: utf-8 -*-

# ref: https://github.com/xmindltd/xmind-sdk-python/tree/master/xmind
import xmind
from xmind.core import workbook,saver
from xmind.core.topic import TopicElement
from xmind.core import const
from xmind.core.topic import ChildrenElement
from xmind.core.markerref import MarkerId
workbook = xmind.load('团聚TCcqj.xmind')
sheet = workbook.getPrimarySheet()
rt = sheet.getRootTopic()

topics = rt.getSubTopics(const.TOPIC_ATTACHED)

def getAllTopics(rootTopic, filter=None):
    topics = []
    queue = [rootTopic]
    index = 0;
    while(index < len(queue)):
        currentTopic = queue[index]
        index += 1
        if filter:
            if filter(currentTopic):
                topics.append(currentTopic)
        else:
            topics.append(currentTopic)
        subTopics = currentTopic.getSubTopics(const.TOPIC_ATTACHED)
        if subTopics:
            for subTopic in subTopics:
                queue.append(subTopic)
    return topics

def filter_marker_id(topic):
    ids = [MarkerId.priority1, MarkerId.priority2, MarkerId.priority3]
    markers = topic.getMarkers()
    if markers:
        for markerRef in markers:
            markerId = markerRef.getMarkerId()
            print(markerId.name)
            if markerRef.getMarkerId().name in ids:
                return True
    return False


topics = getAllTopics(rt, filter_marker_id)
for topic in topics:
    print(topic.getTitle())
    for marker in topic.getMarkers():
        print(marker.getMarkerId())
