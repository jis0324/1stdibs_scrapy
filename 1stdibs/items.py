# -*- coding: utf-8 -*-

# Define here the models for your scraped items
#
# See documentation in:
# https://docs.scrapy.org/en/latest/topics/items.html

import scrapy


class RestaurantItem(scrapy.Item):
    BRAND = scrapy.Field()
    PRODUCT = scrapy.Field()
    NAME = scrapy.Field()
    PRODUCT_TYPE = scrapy.Field()
    QUANTITY = scrapy.Field()
    CONDITION = scrapy.Field()
    DELIVERY_FEE = scrapy.Field()
    ORIGINAL_PRICE = scrapy.Field()
    REQUESTED_PRICE = scrapy.Field()
    SOLD_PRICE = scrapy.Field()
    ITEM_LOCATION = scrapy.Field()
    LIST_DATE = scrapy.Field()
    SOLD_DATE = scrapy.Field()
    LISTING_LINK = scrapy.Field()
