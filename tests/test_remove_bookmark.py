#
# Stickyprints
# Copyright (C) 2017 Maurice van der Pot <griffon26@kfk4ever.com>
#
# This file is part of Stickyprints.
#
# Stickyprints is free software: you can redistribute it and/or modify
# it under the terms of the GNU General Public License as published by
# the Free Software Foundation, either version 3 of the License, or
# (at your option) any later version.
#
# Stickyprints is distributed in the hope that it will be useful,
# but WITHOUT ANY WARRANTY; without even the implied warranty of
# MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
# GNU General Public License for more details.
#
# You should have received a copy of the GNU General Public License
# along with Stickyprints.  If not, see <http://www.gnu.org/licenses/>.
#

import re
import unittest
import xml.etree.ElementTree as et

import stickyprints

def document(xml):
    document_xml = '''
        <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
        <w:body>
        %s
        </w:body>
        </w:document>
    '''

    return document_xml % xml

def taglist(element):
    tag = re.sub(r'{.*}(.*)', r'\1', element.tag)
    if list(element):
        return (tag, [ taglist(el) for el in list(element) ])
    else:
        return tag

class TestRemoveBookmark(unittest.TestCase):

    def setUp(self):
        et.register_namespace('w', stickyprints.WORD_NAMESPACE)

    def test_bookmark_removed_even_when_no_broken_placeholders(self):
        root = et.fromstring(document('''
            <child1/>
            <w:bookmarkStart w:id="0" w:name="_GoBack"/>
            <child2/>
            <w:bookmarkEnd w:id="0"/>
            <child3/>
        '''))

        stickyprints.remove_go_back_bookmark(root)

        self.assertEqual(taglist(root),
            ('document', [
                ('body', [
                    ('child1'),
                    ('child2'),
                    ('child3')
                ])
            ])
        )

    def test_endmark_not_belonging_to_startmark_is_not_removed(self):
        root = et.fromstring(document('''
            <child1/>
            <w:bookmarkStart w:id="0" w:name="_GoBack"/>
            <w:bookmarkEnd w:id="1"/>
            <child2/>
        '''))

        stickyprints.remove_go_back_bookmark(root)

        self.assertEqual(taglist(root),
            ('document', [
                ('body', [
                    ('child1'),
                    ('bookmarkEnd'),
                    ('child2')
                ])
            ])
        )

    def test_placeholder_interrupted_by_bookmark_start_and_end_is_restored(self):
        root = et.fromstring(document('''
            <tag1>textbefore &lt;place</tag1>
            <w:bookmarkStart w:id="0" w:name="_GoBack"/>
            <tag2/>
            <w:bookmarkEnd w:id="0"/>
            <tag3/>
            <tag1>holder&gt; textafter</tag1>
            <tag3/>
        '''))

        stickyprints.remove_go_back_bookmark(root)

        self.assertEqual(taglist(root),
            ('document', [
                ('body', [
                    ('tag1'),
                    ('tag3')
                ])
            ])
        )
        self.assertEqual(root.find('.//tag1').text, 'textbefore <placeholder> textafter')

    def test_placeholder_interrupted_by_only_bookmark_start_is_restored(self):
        root = et.fromstring(document('''
            <tag1>textbefore &lt;place</tag1>
            <w:bookmarkStart w:id="0" w:name="_GoBack"/>
            <tag2/>
            <tag1>holder&gt; textafter</tag1>
            <tag3/>
        '''))

        stickyprints.remove_go_back_bookmark(root)

        self.assertEqual(taglist(root),
            ('document', [
                ('body', [
                    ('tag1'),
                    ('tag3')
                ])
            ])
        )
        self.assertEqual(root.find('.//tag1').text, 'textbefore <placeholder> textafter')

    def test_placeholder_interrupted_by_only_bookmark_end_is_restored(self):
        root = et.fromstring(document('''
            <w:bookmarkStart w:id="0" w:name="_GoBack"/>
            <tag1/>
            <tag2>textbefore &lt;place</tag2>
            <w:bookmarkEnd w:id="0"/>
            <tag2>holder&gt; textafter</tag2>
            <tag3/>
        '''))

        stickyprints.remove_go_back_bookmark(root)

        self.assertEqual(taglist(root),
            ('document', [
                ('body', [
                    ('tag1'),
                    ('tag2'),
                    ('tag3')
                ])
            ])
        )
        self.assertEqual(root.find('.//tag2').text, 'textbefore <placeholder> textafter')

    def test_parts_of_placeholder_in_elements_with_differing_tags_are_not_glued_together(self):
        root = et.fromstring(document('''
            <tag1>textbefore &lt;place</tag1>
            <w:bookmarkStart w:id="0" w:name="_GoBack"/>
            <w:bookmarkEnd w:id="0"/>
            <tag2>holder&gt; textafter</tag2>
        '''))

        stickyprints.remove_go_back_bookmark(root)

        self.assertEqual(taglist(root),
            ('document', [
                ('body', [
                    ('tag1'),
                    ('tag2')
                ])
            ])
        )
        self.assertEqual(root.find('.//tag1').text, 'textbefore <place')
        self.assertEqual(root.find('.//tag2').text, 'holder> textafter')

if __name__ == '__main__':
    unittest.main()

