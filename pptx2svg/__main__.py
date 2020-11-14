#===============================================================================
#
#  Flatmap viewer and annotation tools
#
#  Copyright (c) 2019  David Brooks
#
#  Licensed under the Apache License, Version 2.0 (the "License");
#  you may not use this file except in compliance with the License.
#  You may obtain a copy of the License at
#
#      http://www.apache.org/licenses/LICENSE-2.0
#
#  Unless required by applicable law or agreed to in writing, software
#  distributed under the License is distributed on an "AS IS" BASIS,
#  WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
#  See the License for the specific language governing permissions and
#  limitations under the License.
#
#===============================================================================

from math import sqrt, sin, cos, pi as PI

import os
import re
import string

#===============================================================================

import numpy as np

import pptx.shapes.connector
from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE

import svgwrite

from tqdm import tqdm

#===============================================================================

from formula import Geometry, radians
from presets import DML

#===============================================================================

# Internal PPT units are EMUs (English Metric Units)

EMU_PER_CM  = 360000
EMU_PER_IN  = 914400

DOTS_PER_IN = 96  ### 300 ??

EMU_PER_DOT = EMU_PER_IN/DOTS_PER_IN

#===============================================================================

def transform_point(transform, point):
#=====================================
    return (transform@[point[0], point[1], 1.0])[:2]

#===============================================================================

def ellipse_point(a, b, theta):
#==============================
    a_sin_theta = a*sin(theta)
    b_cos_theta = b*cos(theta)
    circle_radius = sqrt(a_sin_theta**2 + b_cos_theta**2)
    return (a*b_cos_theta/circle_radius, b*a_sin_theta/circle_radius)

#===============================================================================

def match_to_hex(m):
#===================
    c = m[0]
    return (c   if c in (string.ascii_letters + string.digits) else
            '_' if c in string.whitespace else
            '_x{:02X}_'.format(ord(c)))

def adobe_encode(s):
#===================
    return re.sub('.', match_to_hex, s)

#===============================================================================

class DrawMLTransform(object):
    def __init__(self, shape, bbox=None):
        xfrm = shape.element.xfrm

        # From Section L.4.7.6 of ECMA-376 Part 1
        (Bx, By) = ((xfrm.chOff.x, xfrm.chOff.y)
                        if xfrm.chOff is not None else
                    (0, 0))
        (Dx, Dy) = ((xfrm.chExt.cx, xfrm.chExt.cy)
                        if xfrm.chExt is not None else
                    bbox)
        (Bx_, By_) = (xfrm.off.x, xfrm.off.y)
        (Dx_, Dy_) = (xfrm.ext.cx, xfrm.ext.cy)
        theta = xfrm.rot*PI/180.0
        Fx = -1 if xfrm.flipH else 1
        Fy = -1 if xfrm.flipV else 1
        T_st = np.array([[Dx_/Dx,      0, Bx_ - (Dx_/Dx)*Bx] if Dx != 0 else [1, 0, Bx_],
                         [     0, Dy_/Dy, By_ - (Dy_/Dy)*By] if Dy != 0 else [0, 1, By_],
                         [     0,      0,                 1]])
        U = np.array([[1, 0, -(Bx_ + Dx_/2.0)],
                      [0, 1, -(By_ + Dy_/2.0)],
                      [0, 0,                1]])
        R = np.array([[cos(theta), -sin(theta), 0],
                      [sin(theta),  cos(theta), 0],
                      [0,                    0, 1]])
        Flip = np.array([[Fx,  0, 0],
                         [ 0, Fy, 0],
                         [ 0,  0, 1]])
        T_rf = np.linalg.inv(U)@R@Flip@U
        self.__T = T_rf@T_st

    def matrix(self):
        return self.__T

#===============================================================================

class SvgLayer(object):
    def __init__(self, size, slide):
        self.__slide = slide
        self.__dwg = svgwrite.Drawing(filename=None, size=size)
        self.__dwg.defs.add(self.__dwg.style('.non-scaling-stroke { vector-effect: non-scaling-stroke; }'))

    def save(self, filename):
    #========================
        self.__dwg.saveas(filename)

    def process(self, transform):
    #============================
        self.process_shape_list(self.__slide.shapes, self.__dwg, transform, True)

    def process_group(self, group, svg_parent, transform):
    #=====================================================
        svg_group = self.__dwg.g(id=adobe_encode(group.name))
        svg_parent.add(svg_group)
        self.process_shape_list(group.shapes, svg_group, transform@DrawMLTransform(group).matrix())

    def process_shape_list(self, shapes, svg_parent, transform, outermost=False):
    #============================================================================
        if outermost:
            print('Processing shape list...')
            progress_bar = tqdm(total=len(shapes),
                unit='shp', ncols=40,
                bar_format='{l_bar}{bar}| {n_fmt}/{total_fmt}')
        for shape in shapes:
            if (shape.shape_type == MSO_SHAPE_TYPE.AUTO_SHAPE
             or shape.shape_type == MSO_SHAPE_TYPE.FREEFORM
             or isinstance(shape, pptx.shapes.connector.Connector)):
                self.process_shape(shape, svg_parent, transform)
            elif shape.shape_type == MSO_SHAPE_TYPE.GROUP:
                self.process_group(shape, svg_parent, transform)
            elif (shape.shape_type == MSO_SHAPE_TYPE.TEXT_BOX
               or shape.shape_type == MSO_SHAPE_TYPE.PICTURE):
                pass
            else:
                print('"{}" {} not processed...'.format(shape.name, str(shape.shape_type)))
            if outermost:
                progress_bar.update(1)
        if outermost:
            progress_bar.close()

    def process_shape(self, shape, svg_parent, transform):
    #=====================================================
        geometry = Geometry(shape)
        for path in geometry.path_list:
            svg_path = self.__dwg.path(id=adobe_encode(shape.name),
                                       fill='none', class_='non-scaling-stroke')
            bbox = (shape.width, shape.height) if path.w is None else (path.w, path.h)
            T = transform@DrawMLTransform(shape, bbox).matrix()
            first_point = None
            current_point = None
            closed = False
            for c in path.getchildren():
                if   c.tag == DML('arcTo'):
                    wR = geometry.attrib_value(c, 'wR')
                    hR = geometry.attrib_value(c, 'hR')
                    stAng = radians(geometry.attrib_value(c, 'stAng'))
                    swAng = radians(geometry.attrib_value(c, 'swAng'))
                    p1 = ellipse_point(wR, hR, stAng)
                    p2 = ellipse_point(wR, hR, stAng + swAng)
                    pt = (current_point[0] - p1[0] + p2[0],
                          current_point[1] - p1[1] + p2[1])
                    large_arc_flag = 1 if swAng >= PI else 0
                    svg_path.push('A', *transform_point(T, (wR, hR)),
                                       0, large_arc_flag, 1,
                                       *transform_point(T, pt))
                    current_point = pt

                elif c.tag == DML('close'):
                    if first_point is not None and current_point != first_point:
                        svg_path.push('Z')
                    closed = True
                    first_point = None
                elif c.tag == DML('cubicBezTo'):
                    coords = []
                    for p in c.getchildren():
                        pt = geometry.point(p)
                        coords.extend(transform_point(T, pt))
                        current_point = pt
                    svg_path.push('C', *coords)
                elif c.tag == DML('lnTo'):
                    pt = geometry.point(c.pt)
                    coords = transform_point(T, pt)
                    svg_path.push('L', *coords)
                    current_point = pt
                elif c.tag == DML('moveTo'):
                    pt = geometry.point(c.pt)
                    coords = transform_point(T, pt)
                    svg_path.push('M', *coords)
                    if first_point is None:
                        first_point = pt
                    current_point = pt
                elif c.tag == DML('quadBezTo'):
                    coords = []
                    for p in c.getchildren():
                        pt = geometry.point(p)
                        coords.extend(transform_point(T, pt))
                        current_point = pt
                    svg_path.push('Q', *coords)
                else:
                    print('Unknown path element: {}'.format(c.tag))

            ### Get style attributes from PPT
            svg_path.attribs['stroke-width'] = 3
            if closed:
                svg_path.attribs['fill'] = '#808080'   ## shape.fill and shape.line
                svg_path.attribs['opacity'] = 0.3
                svg_path.attribs['stroke'] = 'red'
            else:
                svg_path.attribs['stroke'] = 'blue'

            svg_parent.add(svg_path)

#===============================================================================

class SvgExtractor(object):
    def __init__(self, pptx, output_dir, debug=False):
        self.__pptx = Presentation(pptx)
        self.__slides = self.__pptx.slides
        self.__transform = np.array([[1.0/EMU_PER_DOT,               0, 0],
                                     [              0, 1.0/EMU_PER_DOT, 0],
                                     [              0,               0, 1]])
        (pptx_width, pptx_height) = (self.__pptx.slide_width, self.__pptx.slide_height)
        self.__svg_size = transform_point(self.__transform, (pptx_width, pptx_height))
        self.__output_dir = output_dir
        self.__debug = debug

    def slide_to_svg(self, slide, slide_number):
    #===========================================
        if self.__debug:
            with open(os.path.join(self.__output_dir, 'slide-{:02d}.xml'.format(slide_number)), 'w') as xml:
                xml.write(slide.element.xml)
        layer = SvgLayer(self.__svg_size, slide)
        layer.process(self.__transform)
        layer.save(os.path.join(self.__output_dir, 'slide-{:02d}.svg'.format(slide_number)))

    def slides_to_svg(self):
    #=======================
        for n, slide in enumerate(self.__slides):
            self.slide_to_svg(slide, n+1)

#===============================================================================

if __name__ == '__main__':
    import argparse

    parser = argparse.ArgumentParser(description='Convert Powerpoint slides to SVG.')

    parser.add_argument('-d', '--debug', action='store_true', help='save DrawML to aid with debugging')

    parser.add_argument('powerpoint', metavar='POWERPOINT_FILE',
                        help='the Powerpoint file to convert')

    parser.add_argument('output_dir', metavar='OUTPUT_DIRECTORY',
                        help='directory in which to save SVG files')

    args = parser.parse_args()

    if not os.path.exists(args.output_dir):
        os.makedirs(args.output_dir)

    extractor = SvgExtractor(args.powerpoint, args.output_dir, args.debug)
    extractor.slides_to_svg()

#===============================================================================

